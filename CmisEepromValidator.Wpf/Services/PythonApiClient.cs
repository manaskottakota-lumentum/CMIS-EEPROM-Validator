using System.IO;
using System.Net.Http;
using System.Text.Json;
using CmisEepromValidator.Wpf.Models;

namespace CmisEepromValidator.Wpf.Services;

public sealed class PythonApiClient
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly HttpClient _httpClient;

    public PythonApiClient()
    {
        _httpClient = new HttpClient
        {
            BaseAddress = new Uri("http://127.0.0.1:8765"),
            Timeout = TimeSpan.FromMinutes(2)
        };
    }

    public async Task<ValidationResponse> ValidateAsync(string dumpPath, string workbookPath, CancellationToken cancellationToken)
    {
        await using var dumpStream = OpenSharedRead(dumpPath);
        await using var workbookStream = OpenSharedRead(workbookPath);

        using var content = new MultipartFormDataContent();
        content.Add(new StreamContent(dumpStream), "dump", Path.GetFileName(dumpPath));
        content.Add(new StreamContent(workbookStream), "workbook", Path.GetFileName(workbookPath));

        HttpResponseMessage response;
        try
        {
            response = await _httpClient.PostAsync("/api/validate", content, cancellationToken).ConfigureAwait(false);
        }
        catch (HttpRequestException ex)
        {
            throw new InvalidOperationException("The Python backend could not be reached. Close and reopen the app, or set CMIS_VALIDATOR_PYTHON to a valid python.exe path.", ex);
        }

        var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(ReadError(body));
        }

        return JsonSerializer.Deserialize<ValidationResponse>(body, JsonOptions)
            ?? throw new InvalidOperationException("The Python backend returned an empty validation response.");
    }

    public async Task<ValidationResponse> ValidateSwitchAsync(
        string host,
        string username,
        string password,
        string port,
        string pages,
        string workbookPath,
        CancellationToken cancellationToken)
    {
        var payload = JsonSerializer.Serialize(new
        {
            switchType = "sonic",
            host,
            username,
            password,
            port,
            pages,
            workbookPath
        });

        using var content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json");

        HttpResponseMessage response;
        try
        {
            response = await _httpClient.PostAsync("/api/validate-switch", content, cancellationToken).ConfigureAwait(false);
        }
        catch (HttpRequestException ex)
        {
            throw new InvalidOperationException("The Python backend could not be reached. Close and reopen the app, or set CMIS_VALIDATOR_PYTHON to a valid python.exe path.", ex);
        }

        var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(ReadError(body));
        }

        return JsonSerializer.Deserialize<ValidationResponse>(body, JsonOptions)
            ?? throw new InvalidOperationException("The Python backend returned an empty validation response.");
    }

    private static FileStream OpenSharedRead(string path)
    {
        return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    }

    private static string ReadError(string body)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            return "The Python backend returned an unknown error.";
        }

        try
        {
            using var document = JsonDocument.Parse(body);
            if (document.RootElement.TryGetProperty("error", out var error))
            {
                return error.GetString() ?? body;
            }
        }
        catch (JsonException)
        {
        }

        return body;
    }
}
