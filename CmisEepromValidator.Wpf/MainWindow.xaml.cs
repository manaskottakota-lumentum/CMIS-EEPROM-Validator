using System.Windows;
using CmisEepromValidator.Wpf.Services;
using CmisEepromValidator.Wpf.ViewModels;

namespace CmisEepromValidator.Wpf;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainViewModel(new PythonApiClient(), new FileDialogService());
    }
}
