using System.Globalization;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Windows.Media;

namespace CmisEepromValidator.Wpf.Models;

public sealed class ValidationResponse
{
    [JsonPropertyName("results")]
    public List<ValidationResultDto> Results { get; set; } = [];

    [JsonPropertyName("summary")]
    public ValidationSummary Summary { get; set; } = new();

    [JsonPropertyName("validationTime")]
    public string ValidationTime { get; set; } = "";

    [JsonPropertyName("durationSeconds")]
    public double DurationSeconds { get; set; }
}

public sealed class ValidationSummary
{
    [JsonPropertyName("total")]
    public int Total { get; set; }

    [JsonPropertyName("passed")]
    public int Passed { get; set; }

    [JsonPropertyName("failed")]
    public int Failed { get; set; }

    [JsonPropertyName("read")]
    public int Read { get; set; }

    [JsonPropertyName("notChecked")]
    public int NotChecked { get; set; }

    [JsonPropertyName("errors")]
    public int Errors { get; set; }
}

public sealed class ValidationResultDto
{
    [JsonPropertyName("sourceSheet")]
    public string SourceSheet { get; set; } = "";

    [JsonPropertyName("check")]
    public string Check { get; set; } = "";

    [JsonPropertyName("parameter")]
    public string Parameter { get; set; } = "";

    [JsonPropertyName("description")]
    public string Description { get; set; } = "";

    [JsonPropertyName("page")]
    public string Page { get; set; } = "";

    [JsonPropertyName("bank")]
    public string Bank { get; set; } = "";
    [JsonPropertyName("address")]
    public string Address { get; set; } = "";

    [JsonPropertyName("bits")]
    public string Bits { get; set; } = "";

    [JsonPropertyName("length")]
    public int Length { get; set; }

    [JsonPropertyName("dataType")]
    public string DataType { get; set; } = "";

    [JsonPropertyName("cmisType")]
    public string CmisType { get; set; } = "";

    [JsonPropertyName("simpleType")]
    public string SimpleType { get; set; } = "";

    [JsonPropertyName("expected")]
    public string Expected { get; set; } = "";

    [JsonPropertyName("actual")]
    public string Actual { get; set; } = "";

    [JsonPropertyName("rawHex")]
    public string RawHex { get; set; } = "";

    [JsonPropertyName("status")]
    public string Status { get; set; } = "";

    [JsonPropertyName("message")]
    public string Message { get; set; } = "";

    [JsonPropertyName("activeCheck")]
    public bool ActiveCheck { get; set; }

    [JsonPropertyName("sourceRow")]
    public object? SourceRow { get; set; }
}

public sealed class ValidationRow
{
    private static readonly SolidColorBrush PassBrush = new(Color.FromRgb(4, 120, 87));
    private static readonly SolidColorBrush FailBrush = new(Color.FromRgb(220, 38, 38));
    private static readonly SolidColorBrush NeutralBrush = new(Color.FromRgb(55, 65, 81));

    public ValidationRow(ValidationResultDto dto)
    {
        SourceSheet = NormalizeSourceSheet(dto.SourceSheet);
        Check = string.IsNullOrWhiteSpace(dto.Check) ? "-" : dto.Check;
        Parameter = dto.Parameter;
        Description = dto.Description;
        Page = dto.Page;
        PageDisplay = FormatPageHex(dto.Page);
        PageDecimal = FormatPageDecimal(dto.Page);
        Bank = string.IsNullOrWhiteSpace(dto.Bank) ? "N/A" : dto.Bank;
        Address = dto.Address;
        AddressDisplay = FormatAddressHex(dto.Address);
        AddressDecimal = FormatAddressDecimal(dto.Address);
        PageAddressDecimal = FormatPageAddressDecimal(dto.Page, dto.Address);
        Bits = string.IsNullOrWhiteSpace(dto.Bits) ? "-" : dto.Bits;
        Length = dto.Length;
        DataType = dto.DataType;
        CmisType = dto.CmisType;
        SimpleType = dto.SimpleType;
        Expected = string.IsNullOrWhiteSpace(dto.Expected) ? "-" : dto.Expected;
        Actual = string.IsNullOrWhiteSpace(dto.Actual) ? "-" : dto.Actual;
        RawHex = string.IsNullOrWhiteSpace(dto.RawHex) ? "-" : dto.RawHex;
        Status = dto.Status;
        Message = dto.Message;
        ActiveCheck = dto.ActiveCheck;
        SourceRow = dto.SourceRow?.ToString() ?? "";
    }

    public string SourceSheet { get; }
    public string Check { get; }
    public string Parameter { get; }
    public string Description { get; }
    public string Page { get; }
    public string PageDisplay { get; }
    public string PageDecimal { get; }
    public string Bank { get; }
    public string Address { get; }
    public string AddressDisplay { get; }
    public string AddressDecimal { get; }
    public string PageAddressDecimal { get; }
    public string Bits { get; }
    public int Length { get; }
    public string DataType { get; }
    public string CmisType { get; }
    public string SimpleType { get; }
    public string Expected { get; }
    public string Actual { get; }
    public string RawHex { get; }
    public string Status { get; }
    public string Message { get; }
    public bool ActiveCheck { get; }
    public string SourceRow { get; }

    public string PageAddress => Address;
    public string LengthText => Length > 0 ? $"{Length} byte(s)" : "-";
    public string SearchBlob => $"{SourceSheet} {Check} {Parameter} {Description} {PageDisplay} {PageDecimal} {AddressDisplay} {AddressDecimal} {Bits} {DataType} {Expected} {Actual} {Status}";
    public Brush StatusBrush => Status is "PASS" ? PassBrush : Status is "FAIL" or "ERROR" ? FailBrush : NeutralBrush;

    private static string FormatPageHex(string page)
    {
        if (string.IsNullOrWhiteSpace(page) || page.Equals("lower", StringComparison.OrdinalIgnoreCase))
        {
            return "Lower";
        }

        var normalized = page.Trim().Replace("0x", "", StringComparison.OrdinalIgnoreCase).TrimEnd('h', 'H');
        return int.TryParse(normalized, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value)
            ? $"{value:X2}h"
            : page;
    }

    private static string FormatPageDecimal(string page)
    {
        if (string.IsNullOrWhiteSpace(page) || page.Equals("lower", StringComparison.OrdinalIgnoreCase))
        {
            return "Lower";
        }

        var normalized = page.Trim().Replace("0x", "", StringComparison.OrdinalIgnoreCase).TrimEnd('h', 'H');
        return int.TryParse(normalized, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value)
            ? value.ToString(CultureInfo.InvariantCulture)
            : page;
    }

    private static string FormatAddressHex(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
        {
            return "-";
        }

        return address.Contains('/')
            ? address.Split('/', 2)[1].Trim()
            : address.Trim();
    }
    private static string FormatPageAddressDecimal(string page, string address)
    {
        var pageDisplay = NormalizePageDecimal(page);
        var addressDisplay = FormatAddressDecimal(address);
        return string.IsNullOrWhiteSpace(addressDisplay) ? pageDisplay : $"{pageDisplay} / {addressDisplay}";
    }

    private static string NormalizePageDecimal(string page)
    {
        if (string.IsNullOrWhiteSpace(page) || page.Equals("lower", StringComparison.OrdinalIgnoreCase))
        {
            return "Lower";
        }

        var normalized = page.Trim().Replace("0x", "", StringComparison.OrdinalIgnoreCase).TrimEnd('h', 'H');
        return int.TryParse(normalized, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value)
            ? value.ToString(CultureInfo.InvariantCulture)
            : page;
    }

    private static string FormatAddressDecimal(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
        {
            return "-";
        }

        var addressPart = address.Contains('/')
            ? address.Split('/', 2)[1]
            : address;

        var matches = Regex.Matches(addressPart, @"(?:0x)?([0-9A-Fa-f]+)h?");
        if (matches.Count == 0)
        {
            return addressPart.Trim();
        }

        var values = matches
            .Select(match => int.Parse(match.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture))
            .ToArray();
        return string.Join("-", values);
    }

    private static string NormalizeSourceSheet(string sourceSheet)
    {
        if (string.IsNullOrWhiteSpace(sourceSheet))
        {
            return "-";
        }

        return sourceSheet.Replace(" Memory", "", StringComparison.OrdinalIgnoreCase);
    }
}
