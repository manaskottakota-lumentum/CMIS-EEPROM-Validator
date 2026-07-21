using System.Windows;
using Microsoft.Win32;

namespace CmisEepromValidator.Wpf.Services;

public interface IFileDialogService
{
    string? OpenFile(string title, string filter);
    string? SaveFile(string title, string filter, string defaultFileName);
    void ShowError(string message);
    void ShowInfo(string title, string message);
}

public sealed class FileDialogService : IFileDialogService
{
    public string? OpenFile(string title, string filter)
    {
        var dialog = new OpenFileDialog
        {
            Title = title,
            Filter = filter,
            CheckFileExists = true,
            Multiselect = false
        };
        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    public string? SaveFile(string title, string filter, string defaultFileName)
    {
        var dialog = new SaveFileDialog
        {
            Title = title,
            Filter = filter,
            FileName = defaultFileName,
            AddExtension = true,
            OverwritePrompt = true
        };
        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    public void ShowError(string message)
    {
        MessageBox.Show(message, "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    public void ShowInfo(string title, string message)
    {
        MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information);
    }
}
