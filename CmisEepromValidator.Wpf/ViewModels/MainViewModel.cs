using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using CmisEepromValidator.Wpf.Models;
using CmisEepromValidator.Wpf.Services;

namespace CmisEepromValidator.Wpf.ViewModels;

public sealed class MainViewModel : ViewModelBase
{
    private static readonly Brush InkBrush = new SolidColorBrush(Color.FromRgb(7, 29, 54));
    private static readonly Brush MutedBrush = new SolidColorBrush(Color.FromRgb(95, 107, 122));
    private static readonly Brush PassBrush = new SolidColorBrush(Color.FromRgb(4, 120, 87));
    private static readonly Brush FailBrush = new SolidColorBrush(Color.FromRgb(220, 38, 38));
    private static readonly Brush PassBackgroundBrush = new SolidColorBrush(Color.FromRgb(232, 247, 239));
    private static readonly Brush FailBackgroundBrush = new SolidColorBrush(Color.FromRgb(255, 237, 237));
    private static readonly Brush NeutralBackgroundBrush = new SolidColorBrush(Color.FromRgb(238, 242, 247));

    private readonly PythonApiClient _apiClient;
    private readonly IFileDialogService _dialogs;
    private readonly PythonBackendManager _backendManager;
    private string _dumpPath = "";
    private string _dumpSourceOverride = "";
    private string _workbookPath = "";
    private string _switchHost = "";
    private string _switchUsername = "";
    private string _switchPassword = "";
    private string _switchPort = "";
    private string _switchPages = "";
    private string _searchText = "";
    private string _selectedStatus = "All Status";
    private bool _showAllRows;
    private bool _isBusy;
    private ValidationRow? _selectedResult;
    private string _overallText = "Not run";
    private string _overallMessage = "Load files and run validation.";
    private Brush _overallBrush = InkBrush;
    private int _total;
    private int _passed;
    private int _failed;
    private int _read;
    private int _notChecked;
    private int _errors;
    private string _validationTime = "-";
    private string _durationText = "Duration: -";
    private string _footerStatus = "Ready";
    private bool _isResultsExpanded;

    public MainViewModel(PythonApiClient apiClient, IFileDialogService dialogs)
        : this(apiClient, dialogs, new PythonBackendManager())
    {
    }

    public MainViewModel(PythonApiClient apiClient, IFileDialogService dialogs, PythonBackendManager backendManager)
    {
        _apiClient = apiClient;
        _dialogs = dialogs;
        _backendManager = backendManager;
        ResultsView = CollectionViewSource.GetDefaultView(Results);
        ResultsView.Filter = FilterResult;

        BrowseDumpCommand = new RelayCommand(_ => BrowseDump());
        BrowseWorkbookCommand = new RelayCommand(_ => BrowseWorkbook());
        ViewDumpInfoCommand = new RelayCommand(_ => ShowFileInfo("EEPROM / CMIS Hex Dump", DumpPath), _ => File.Exists(DumpPath));
        ViewWorkbookInfoCommand = new RelayCommand(_ => ShowFileInfo("Expected Values", WorkbookPath), _ => File.Exists(WorkbookPath));
        LoadSamplesCommand = new RelayCommand(_ => LoadSamples());
        ClearCommand = new RelayCommand(_ => Clear());
        ExportResultsCommand = new RelayCommand(_ => ExportResults());
        ToggleResultsExpandedCommand = new RelayCommand(_ => IsResultsExpanded = !IsResultsExpanded);
        RunValidationCommand = new AsyncRelayCommand(RunValidationAsync, CanRunValidation);
        ReadSwitchCommand = new AsyncRelayCommand(ReadSwitchAsync, CanReadSwitch);
    }

    public ObservableCollection<ValidationRow> Results { get; } = [];
    public ICollectionView ResultsView { get; }
    public IReadOnlyList<string> StatusOptions { get; } = ["All Status", "PASS", "FAIL", "READ", "NOT CHECKED", "ERROR"];

    public RelayCommand BrowseDumpCommand { get; }
    public RelayCommand BrowseWorkbookCommand { get; }
    public RelayCommand ViewDumpInfoCommand { get; }
    public RelayCommand ViewWorkbookInfoCommand { get; }
    public RelayCommand LoadSamplesCommand { get; }
    public RelayCommand ClearCommand { get; }
    public RelayCommand ExportResultsCommand { get; }
    public RelayCommand ToggleResultsExpandedCommand { get; }
    public AsyncRelayCommand RunValidationCommand { get; }
    public AsyncRelayCommand ReadSwitchCommand { get; }

    public string DumpPath
    {
        get => _dumpPath;
        set
        {
            if (SetProperty(ref _dumpPath, value))
            {
                _dumpSourceOverride = "";
                OnFileChanged(nameof(DumpFileName), nameof(DumpFileSize), nameof(DumpLoadedText), nameof(DumpLoadedBrush));
            }
        }
    }

    public string SwitchHost
    {
        get => _switchHost;
        set
        {
            if (SetProperty(ref _switchHost, value))
            {
                ReadSwitchCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public string SwitchUsername
    {
        get => _switchUsername;
        set
        {
            if (SetProperty(ref _switchUsername, value))
            {
                ReadSwitchCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public string SwitchPassword
    {
        get => _switchPassword;
        set => SetProperty(ref _switchPassword, value);
    }

    public string SwitchPort
    {
        get => _switchPort;
        set
        {
            if (SetProperty(ref _switchPort, value))
            {
                ReadSwitchCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public string SwitchPages
    {
        get => _switchPages;
        set => SetProperty(ref _switchPages, value);
    }

    public string WorkbookPath
    {
        get => _workbookPath;
        set
        {
            if (SetProperty(ref _workbookPath, value))
            {
                OnFileChanged(nameof(WorkbookFileName), nameof(WorkbookFileSize), nameof(WorkbookLoadedText), nameof(WorkbookLoadedBrush));
            }
        }
    }

    public string SearchText
    {
        get => _searchText;
        set
        {
            if (SetProperty(ref _searchText, value))
            {
                ResultsView.Refresh();
                OnPropertyChanged(nameof(SearchPlaceholderVisibility));
                OnPropertyChanged(nameof(ResultRangeText));
            }
        }
    }

    public string SelectedStatus
    {
        get => _selectedStatus;
        set
        {
            if (SetProperty(ref _selectedStatus, value))
            {
                ResultsView.Refresh();
                OnPropertyChanged(nameof(ResultRangeText));
            }
        }
    }

    public bool ShowAllRows
    {
        get => _showAllRows;
        set
        {
            if (SetProperty(ref _showAllRows, value))
            {
                ResultsView.Refresh();
                UpdateSummary();
                OnPropertyChanged(nameof(ResultRangeText));
                SelectedResult = ResultsView.Cast<ValidationRow>().FirstOrDefault();
            }
        }
    }

    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            if (SetProperty(ref _isBusy, value))
            {
                RunValidationCommand.RaiseCanExecuteChanged();
                ReadSwitchCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public ValidationRow? SelectedResult
    {
        get => _selectedResult;
        set => SetProperty(ref _selectedResult, value);
    }

    public string OverallText
    {
        get => _overallText;
        private set => SetProperty(ref _overallText, value);
    }

    public string OverallMessage
    {
        get => _overallMessage;
        private set => SetProperty(ref _overallMessage, value);
    }

    public Brush OverallBrush
    {
        get => _overallBrush;
        private set => SetProperty(ref _overallBrush, value);
    }

    public int Total
    {
        get => _total;
        private set => SetProperty(ref _total, value);
    }

    public int Passed
    {
        get => _passed;
        private set => SetProperty(ref _passed, value);
    }

    public int Failed
    {
        get => _failed;
        private set => SetProperty(ref _failed, value);
    }

    public int Read
    {
        get => _read;
        private set => SetProperty(ref _read, value);
    }

    public int NotChecked
    {
        get => _notChecked;
        private set => SetProperty(ref _notChecked, value);
    }

    public int Errors
    {
        get => _errors;
        private set => SetProperty(ref _errors, value);
    }

    public string ValidationTime
    {
        get => _validationTime;
        private set => SetProperty(ref _validationTime, value);
    }

    public string DurationText
    {
        get => _durationText;
        private set => SetProperty(ref _durationText, value);
    }

    public string FooterStatus
    {
        get => _footerStatus;
        private set => SetProperty(ref _footerStatus, value);
    }

    public bool IsResultsExpanded
    {
        get => _isResultsExpanded;
        set
        {
            if (SetProperty(ref _isResultsExpanded, value))
            {
                OnPropertyChanged(nameof(ResultsColumnSpan));
                OnPropertyChanged(nameof(ResultsAreaGridRow));
                OnPropertyChanged(nameof(ResultsAreaGridRowSpan));
                OnPropertyChanged(nameof(ResultsAreaZIndex));
                OnPropertyChanged(nameof(ResultsAreaMargin));
                OnPropertyChanged(nameof(DetailsVisibility));
                OnPropertyChanged(nameof(ResultsExpandButtonText));
                OnPropertyChanged(nameof(ResultsExpandIconData));
            }
        }
    }

    public string SummaryIconText => OverallText == "PASS" ? "✓" : OverallText == "FAIL" ? "!" : "-";
    public Brush SummaryIconBrush => OverallText == "PASS" ? PassBrush : OverallText == "FAIL" ? FailBrush : InkBrush;
    public Brush SummaryIconBackground => OverallText == "PASS" ? PassBackgroundBrush : OverallText == "FAIL" ? FailBackgroundBrush : NeutralBackgroundBrush;
    public Visibility SearchPlaceholderVisibility => string.IsNullOrWhiteSpace(SearchText) ? Visibility.Visible : Visibility.Collapsed;
    public int ResultsColumnSpan => IsResultsExpanded ? 3 : 1;
    public int ResultsAreaGridRow => IsResultsExpanded ? 1 : 4;
    public int ResultsAreaGridRowSpan => IsResultsExpanded ? 4 : 1;
    public int ResultsAreaZIndex => IsResultsExpanded ? 10 : 0;
    public Thickness ResultsAreaMargin => IsResultsExpanded ? new Thickness(0, 22, 0, 0) : new Thickness(0, 16, 0, 0);
    public Visibility DetailsVisibility => IsResultsExpanded ? Visibility.Collapsed : Visibility.Visible;
    public string ResultsExpandButtonText => IsResultsExpanded ? "Collapse" : "Expand";
    public string ResultsExpandIconData => IsResultsExpanded
        ? "M 4 8 L 8 8 L 8 4 M 16 8 L 12 8 L 12 4 M 4 12 L 8 12 L 8 16 M 16 12 L 12 12 L 12 16"
        : "M 4 8 L 4 4 L 8 4 M 16 8 L 16 4 L 12 4 M 4 12 L 4 16 L 8 16 M 16 12 L 16 16 L 12 16";
    public string ResultRangeText
    {
        get
        {
            var count = ResultsView.Cast<ValidationRow>().Count();
            return count == 0 ? "Showing 0 parameters" : $"Showing 1 to {count} of {count} parameters";
        }
    }

    public string DumpFileName => !string.IsNullOrWhiteSpace(_dumpSourceOverride) ? _dumpSourceOverride : File.Exists(DumpPath) ? Path.GetFileName(DumpPath) : "No dump selected";
    public string DumpFileSize => !string.IsNullOrWhiteSpace(_dumpSourceOverride) ? "Switch" : FormatFileSize(DumpPath);
    public string DumpLoadedText => !string.IsNullOrWhiteSpace(_dumpSourceOverride) ? "Read" : File.Exists(DumpPath) ? "Loaded" : "Not loaded";
    public Brush DumpLoadedBrush => !string.IsNullOrWhiteSpace(_dumpSourceOverride) || File.Exists(DumpPath) ? PassBrush : MutedBrush;

    public string WorkbookFileName => File.Exists(WorkbookPath) ? Path.GetFileName(WorkbookPath) : "No expected-values file selected";
    public string WorkbookFileSize => FormatFileSize(WorkbookPath);
    public string WorkbookLoadedText => File.Exists(WorkbookPath) ? "Loaded" : "Not loaded";
    public Brush WorkbookLoadedBrush => File.Exists(WorkbookPath) ? PassBrush : MutedBrush;

    private void BrowseDump()
    {
        var path = _dialogs.OpenFile("Select EEPROM / CMIS hex dump", "Hex/Text Files (*.hex;*.txt;*.dump;*.log)|*.hex;*.txt;*.dump;*.log|All Files (*.*)|*.*");
        if (!string.IsNullOrWhiteSpace(path))
        {
            DumpPath = path;
        }
    }

    private void BrowseWorkbook()
    {
        var path = _dialogs.OpenFile("Select expected-values workbook", "Spreadsheet Files (*.xlsx;*.csv;*.tsv;*.txt)|*.xlsx;*.csv;*.tsv;*.txt|All Files (*.*)|*.*");
        if (!string.IsNullOrWhiteSpace(path))
        {
            WorkbookPath = path;
        }
    }

    public async Task InitializeBackendAsync()
    {
        try
        {
            FooterStatus = "Starting backend...";
            await _backendManager.EnsureRunningAsync();
            FooterStatus = "Ready";
        }
        catch (Exception ex)
        {
            FooterStatus = "Backend failed";
            _dialogs.ShowError(ex.Message);
        }
    }

    private async Task RunValidationAsync()
    {
        try
        {
            IsBusy = true;
            FooterStatus = "Checking backend...";
            await _backendManager.EnsureRunningAsync();
            FooterStatus = "Running validation...";
            var response = await _apiClient.ValidateAsync(DumpPath, WorkbookPath, CancellationToken.None);
            ApplyResponse(response);
            FooterStatus = "Ready";
        }
        catch (Exception ex)
        {
            FooterStatus = "Validation failed";
            _dialogs.ShowError(ex.Message);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private bool CanRunValidation()
    {
        return !IsBusy && File.Exists(DumpPath) && File.Exists(WorkbookPath);
    }

    private async Task ReadSwitchAsync()
    {
        try
        {
            IsBusy = true;
            FooterStatus = "Checking backend...";
            await _backendManager.EnsureRunningAsync();
            FooterStatus = $"Reading {SwitchPort} from switch...";
            var response = await _apiClient.ValidateSwitchAsync(
                SwitchHost.Trim(),
                SwitchUsername.Trim(),
                SwitchPassword,
                SwitchPort.Trim(),
                SwitchPages.Trim(),
                WorkbookPath,
                CancellationToken.None);
            _dumpSourceOverride = $"{SwitchHost.Trim()} / {SwitchPort.Trim()}";
            OnFileChanged(nameof(DumpFileName), nameof(DumpFileSize), nameof(DumpLoadedText), nameof(DumpLoadedBrush));
            ApplyResponse(response);
            FooterStatus = "Ready";
        }
        catch (Exception ex)
        {
            FooterStatus = "Switch read failed";
            _dialogs.ShowError(ex.Message);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private bool CanReadSwitch()
    {
        return !IsBusy
            && File.Exists(WorkbookPath)
            && !string.IsNullOrWhiteSpace(SwitchHost)
            && !string.IsNullOrWhiteSpace(SwitchUsername)
            && !string.IsNullOrWhiteSpace(SwitchPort);
    }

    private void ApplyResponse(ValidationResponse response)
    {
        Results.Clear();
        foreach (var dto in response.Results)
        {
            Results.Add(new ValidationRow(dto));
        }

        ValidationTime = string.IsNullOrWhiteSpace(response.ValidationTime) ? "-" : response.ValidationTime;
        DurationText = $"Duration: {TimeSpan.FromSeconds(response.DurationSeconds):hh\\:mm\\:ss}";
        ResultsView.Refresh();
        UpdateSummary();
        OnPropertyChanged(nameof(ResultRangeText));
        SelectedResult = ResultsView.Cast<ValidationRow>().FirstOrDefault() ?? Results.FirstOrDefault();
        ExportResultsCommand.RaiseCanExecuteChanged();
    }

    private void UpdateSummary()
    {
        var scopedResults = Results.Where(row => ShowAllRows || row.ActiveCheck).ToList();
        Total = scopedResults.Count;
        Passed = scopedResults.Count(row => row.Status == "PASS");
        Failed = scopedResults.Count(row => row.Status == "FAIL");
        Read = scopedResults.Count(row => row.Status == "READ");
        NotChecked = scopedResults.Count(row => row.Status == "NOT CHECKED");
        Errors = scopedResults.Count(row => row.Status == "ERROR");

        if (Total == 0)
        {
            OverallText = "Not run";
            OverallMessage = "Load files and run validation.";
            OverallBrush = InkBrush;
        }
        else if (Failed > 0 || Errors > 0)
        {
            OverallText = "FAIL";
            OverallMessage = "One or more active checks did not match the expected value.";
            OverallBrush = FailBrush;
        }
        else if (Passed > 0)
        {
            OverallText = "PASS";
            OverallMessage = "All active checks matched the expected values.";
            OverallBrush = PassBrush;
        }
        else if (Read > 0)
        {
            OverallText = "READ";
            OverallMessage = "Rows were read, but no expected values were supplied for pass/fail comparison.";
            OverallBrush = InkBrush;
        }
        else
        {
            OverallText = "NOT CHECKED";
            OverallMessage = "No active checks are selected in the workbook.";
            OverallBrush = InkBrush;
        }

        OnPropertyChanged(nameof(SummaryIconText));
        OnPropertyChanged(nameof(SummaryIconBrush));
        OnPropertyChanged(nameof(SummaryIconBackground));
        OnPropertyChanged(nameof(ResultRangeText));
    }

    private bool FilterResult(object item)
    {
        if (item is not ValidationRow row)
        {
            return false;
        }

        if (!ShowAllRows && !row.ActiveCheck)
        {
            return false;
        }

        if (SelectedStatus != "All Status" && row.Status != SelectedStatus)
        {
            return false;
        }

        return string.IsNullOrWhiteSpace(SearchText)
            || row.SearchBlob.Contains(SearchText, StringComparison.OrdinalIgnoreCase);
    }

    private void LoadSamples()
    {
        var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var documentsRoot = Path.Combine(userProfile, "OneDrive - Lumentum Operations LLC", "Documents");
        var dump = Path.Combine(documentsRoot, "GUI", "outputs", "cmis_lower_memory_overview", "msft_1p6t_2xdr4_validator_test_dump.hex");
        var workbook = Path.Combine(documentsRoot, "1_6T_2xDR4_MemoryMap_MSFT_requirements_v1_2.xlsx");

        if (!File.Exists(dump) || !File.Exists(workbook))
        {
            _dialogs.ShowError("Could not find the MSFT sample dump and workbook files.");
            return;
        }

        DumpPath = dump;
        WorkbookPath = workbook;
        FooterStatus = "MSFT sample files loaded";
    }

    private void Clear()
    {
        Results.Clear();
        _dumpSourceOverride = "";
        OnFileChanged(nameof(DumpFileName), nameof(DumpFileSize), nameof(DumpLoadedText), nameof(DumpLoadedBrush));
        SearchText = "";
        SelectedStatus = "All Status";
        SelectedResult = null;
        ValidationTime = "-";
        DurationText = "Duration: -";
        FooterStatus = "Ready";
        UpdateSummary();
        OnPropertyChanged(nameof(ResultRangeText));
        ExportResultsCommand.RaiseCanExecuteChanged();
    }

    private void ExportResults()
    {
        if (Results.Count == 0)
        {
            _dialogs.ShowError("Run validation before exporting results.");
            return;
        }

        var path = _dialogs.SaveFile("Export validation results", "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*", "cmis_validation_results.csv");
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        var builder = new StringBuilder();
        builder.AppendLine("source_sheet,check,parameter,page,address,data_type,expected,module_value,status,raw_hex,significance");
        foreach (var row in Results)
        {
            builder.AppendLine(string.Join(
                ",",
                Csv(row.SourceSheet),
                Csv(row.Check),
                Csv(row.Parameter),
                Csv(row.Page),
                Csv(row.Address),
                Csv(row.DataType),
                Csv(row.Expected),
                Csv(row.Actual),
                Csv(row.Status),
                Csv(row.RawHex),
                Csv(row.Message)));
        }

        File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        FooterStatus = "Results exported";
    }

    private void ShowFileInfo(string title, string path)
    {
        if (!File.Exists(path))
        {
            return;
        }

        var info = new FileInfo(path);
        _dialogs.ShowInfo(title, $"{info.FullName}{Environment.NewLine}{Environment.NewLine}Size: {FormatFileSize(info.FullName)}");
    }

    private void OnFileChanged(params string[] propertyNames)
    {
        foreach (var propertyName in propertyNames)
        {
            OnPropertyChanged(propertyName);
        }

        RunValidationCommand.RaiseCanExecuteChanged();
        ReadSwitchCommand.RaiseCanExecuteChanged();
        ViewDumpInfoCommand.RaiseCanExecuteChanged();
        ViewWorkbookInfoCommand.RaiseCanExecuteChanged();
    }

    private static string FindProjectRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "api_server.py")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        return Directory.GetCurrentDirectory();
    }

    private static string FormatFileSize(string path)
    {
        if (!File.Exists(path))
        {
            return "-";
        }

        var length = new FileInfo(path).Length;
        if (length >= 1024 * 1024)
        {
            return $"{length / 1024d / 1024d:0.#} MB";
        }

        if (length >= 1024)
        {
            return $"{length / 1024d:0.#} KB";
        }

        return $"{length} B";
    }

    private static string Csv(string value)
    {
        var escaped = value.Replace("\"", "\"\"");
        return escaped.Contains(',') || escaped.Contains('"') || escaped.Contains('\n') || escaped.Contains('\r')
            ? $"\"{escaped}\""
            : escaped;
    }
}
