using System.Globalization;
using System.Text;
using System.Text.Json;

namespace M365Backup;

internal static class IcsWriter
{
    public static async Task WriteEventAsync(
        string path,
        JsonElement calendarEvent,
        CancellationToken cancellationToken)
    {
        var lines = new List<string>
        {
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//M365Backup//EN",
            "CALSCALE:GREGORIAN",
            "BEGIN:VEVENT"
        };

        var id = GetString(calendarEvent, "iCalUId") ?? GetString(calendarEvent, "id")
            ?? throw new InvalidOperationException("Microsoft Graph event is missing an identifier.");
        lines.Add($"UID:{Escape(id)}");
        lines.Add($"DTSTAMP:{DateTime.UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture)}");
        AddEventDate(lines, calendarEvent, "start", "DTSTART");
        AddEventDate(lines, calendarEvent, "end", "DTEND");
        lines.Add($"SUMMARY:{Escape(GetString(calendarEvent, "subject") ?? "no-subject")}");

        var body = GetNestedString(calendarEvent, "body", "content");
        if (!string.IsNullOrWhiteSpace(body))
        {
            lines.Add($"DESCRIPTION:{Escape(body)}");
        }

        var location = GetNestedString(calendarEvent, "location", "displayName");
        if (!string.IsNullOrWhiteSpace(location))
        {
            lines.Add($"LOCATION:{Escape(location)}");
        }

        var organizer = GetNestedString(calendarEvent, "organizer", "emailAddress", "address");
        if (!string.IsNullOrWhiteSpace(organizer))
        {
            lines.Add($"ORGANIZER:mailto:{Escape(organizer)}");
        }

        if (calendarEvent.TryGetProperty("attendees", out var attendees) &&
            attendees.ValueKind == JsonValueKind.Array)
        {
            foreach (var attendee in attendees.EnumerateArray())
            {
                var address = GetNestedString(attendee, "emailAddress", "address");
                if (!string.IsNullOrWhiteSpace(address))
                {
                    lines.Add($"ATTENDEE:mailto:{Escape(address)}");
                }
            }
        }

        lines.Add("END:VEVENT");
        lines.Add("END:VCALENDAR");

        var content = new StringBuilder();
        foreach (var line in lines)
        {
            AppendFoldedLine(content, line);
        }

        await File.WriteAllTextAsync(path, content.ToString(), new UTF8Encoding(false), cancellationToken);
    }

    private static void AddEventDate(List<string> lines, JsonElement calendarEvent, string property, string name)
    {
        if (!calendarEvent.TryGetProperty(property, out var dateValue) ||
            !dateValue.TryGetProperty("dateTime", out var dateTimeValue) ||
            dateTimeValue.ValueKind != JsonValueKind.String ||
            !DateTime.TryParse(
                dateTimeValue.GetString(),
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal,
                out var dateTime))
        {
            throw new InvalidOperationException($"Event does not contain a valid '{property}' date.");
        }

        if (GetBoolean(calendarEvent, "isAllDay"))
        {
            lines.Add($"{name};VALUE=DATE:{dateTime.ToString("yyyyMMdd", CultureInfo.InvariantCulture)}");
        }
        else
        {
            lines.Add($"{name}:{dateTime.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture)}");
        }
    }

    private static string Escape(string value) => value
        .Replace("\\", "\\\\", StringComparison.Ordinal)
        .Replace("\r\n", "\\n", StringComparison.Ordinal)
        .Replace("\n", "\\n", StringComparison.Ordinal)
        .Replace("\r", "\\n", StringComparison.Ordinal)
        .Replace(",", "\\,", StringComparison.Ordinal)
        .Replace(";", "\\;", StringComparison.Ordinal);

    private static string? GetString(JsonElement element, string property) =>
        element.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    private static string? GetNestedString(JsonElement element, params string[] properties)
    {
        foreach (var property in properties)
        {
            if (!element.TryGetProperty(property, out element))
            {
                return null;
            }
        }

        return element.ValueKind == JsonValueKind.String ? element.GetString() : null;
    }

    private static bool GetBoolean(JsonElement element, string property) =>
        element.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.True;

    private static void AppendFoldedLine(StringBuilder output, string line)
    {
        var currentBytes = 0;
        foreach (var rune in line.EnumerateRunes())
        {
            var text = rune.ToString();
            var runeBytes = Encoding.UTF8.GetByteCount(text);
            if (currentBytes + runeBytes > 75)
            {
                output.Append("\r\n ");
                currentBytes = 1;
            }

            output.Append(text);
            currentBytes += runeBytes;
        }

        output.Append("\r\n");
    }
}
