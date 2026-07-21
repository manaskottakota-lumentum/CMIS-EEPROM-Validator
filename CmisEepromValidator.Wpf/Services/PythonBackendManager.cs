using System.Diagnostics;
using System.IO;
using System.Net.Http;

namespace CmisEepromValidator.Wpf.Services;

public sealed class PythonBackendManager : IDisposable
{
    private const string HealthUrl = "http://127.0.0.1:8765/api/health";
    private Process? _backendProcess;

    public async Task EnsureRunningAsync(CancellationToken cancellationToken = default)
    {
        if (await IsRunningAsync(cancellationToken).ConfigureAwait(false))
        {
            return;
        }

        var root = FindProjectRoot();
        var apiServerPath = Path.Combine(root, "api_server.py");
        if (!File.Exists(apiServerPath))
        {
            throw new InvalidOperationException("Could not find api_server.py for the Python backend.");
        }

        var python = FindPythonExecutable(root);
        var arguments = python.FileName.Equals("py", StringComparison.OrdinalIgnoreCase)
            ? $"-3 {Quote(apiServerPath)}"
            : Quote(apiServerPath);

        var startInfo = new ProcessStartInfo
        {
            FileName = python.FileName,
            Arguments = arguments,
            WorkingDirectory = root,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        try
        {
            _backendProcess = Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Could not start the Python backend using {python.DisplayName}. Set CMIS_VALIDATOR_PYTHON to a valid python.exe path if needed.",
                ex);
        }

        if (_backendProcess is null)
        {
            throw new InvalidOperationException("Could not start the Python backend process.");
        }

        await WaitForStartupAsync(cancellationToken).ConfigureAwait(false);
    }

    public void Dispose()
    {
        if (_backendProcess is { HasExited: false })
        {
            try
            {
                _backendProcess.Kill(entireProcessTree: true);
            }
            catch
            {
            }
        }

        _backendProcess?.Dispose();
    }

    private static async Task<bool> IsRunningAsync(CancellationToken cancellationToken)
    {
        using var client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(1)
        };

        try
        {
            using var response = await client.GetAsync(HealthUrl, cancellationToken).ConfigureAwait(false);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    private async Task WaitForStartupAsync(CancellationToken cancellationToken)
    {
        for (var attempt = 0; attempt < 40; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (_backendProcess is { HasExited: true })
            {
                throw new InvalidOperationException("The Python backend exited before it became ready.");
            }

            if (await IsRunningAsync(cancellationToken).ConfigureAwait(false))
            {
                return;
            }

            await Task.Delay(250, cancellationToken).ConfigureAwait(false);
        }

        throw new InvalidOperationException("The Python backend did not become ready in time.");
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

        throw new InvalidOperationException("Could not locate the CMIS-EEPROM-Validator project root.");
    }

    private static PythonExecutable FindPythonExecutable(string root)
    {
        var envPath = Environment.GetEnvironmentVariable("CMIS_VALIDATOR_PYTHON");
        if (!string.IsNullOrWhiteSpace(envPath) && File.Exists(envPath))
        {
            return new PythonExecutable(envPath, envPath);
        }

        var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var candidates = new[]
        {
            Path.Combine(root, ".venv", "Scripts", "python.exe"),
            Path.Combine(root, "venv", "Scripts", "python.exe"),
            Path.Combine(userProfile, ".cache", "codex-runtimes", "codex-primary-runtime", "dependencies", "python", "python.exe")
        };

        foreach (var candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return new PythonExecutable(candidate, candidate);
            }
        }

        return new PythonExecutable("py", "Python launcher");
    }

    private static string Quote(string value)
    {
        return $"\"{value.Replace("\"", "\\\"")}\"";
    }

    private readonly record struct PythonExecutable(string FileName, string DisplayName);
}
