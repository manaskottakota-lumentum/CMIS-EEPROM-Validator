using System.Windows;
using CmisEepromValidator.Wpf.Services;
using CmisEepromValidator.Wpf.ViewModels;

namespace CmisEepromValidator.Wpf;

public partial class ModernWindow : Window
{
    private readonly PythonBackendManager _backendManager = new();
    private readonly MainViewModel _viewModel;

    public ModernWindow()
    {
        InitializeComponent();
        _viewModel = new MainViewModel(new PythonApiClient(), new FileDialogService(), _backendManager);
        DataContext = _viewModel;
        _viewModel.PropertyChanged += OnViewModelPropertyChanged;
        UpdateExpandedColumnsVisibility();
        Loaded += OnLoaded;
        Closed += OnClosed;
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        await _viewModel.InitializeBackendAsync();
    }

    private void OnClosed(object? sender, EventArgs e)
    {
        _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
        _backendManager.Dispose();
    }

    private void OnViewModelPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MainViewModel.IsResultsExpanded))
        {
            UpdateExpandedColumnsVisibility();
        }
    }

    private void UpdateExpandedColumnsVisibility()
    {
        var visibility = _viewModel.IsResultsExpanded ? Visibility.Visible : Visibility.Collapsed;
        AddressDecimalColumn.Visibility = visibility;
        BitsColumn.Visibility = visibility;
    }
}
