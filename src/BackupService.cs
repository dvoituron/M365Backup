using System.Globalization;
using System.Text.Json;

namespace M365Backup;

internal sealed class BackupService(GraphReadClient graph)
{
    public async Task<BackupResult> RunAsync(
        CliOptions options,
        string emailAddress,
        CancellationToken cancellationToken)
    {
        var now = DateTimeOffset.UtcNow;
        var cutoff = now.AddDays(-options.Days);
        var outputDirectory = Path.GetFullPath(options.OutputDirectory);
        if (File.Exists(outputDirectory))
        {
            throw new IOException($"The destination path exists but is not a directory: {outputDirectory}");
        }

        var emailDirectory = Path.Combine(outputDirectory, "email");
        var inboxDirectory = Path.Combine(emailDirectory, "inbox");
        var sentDirectory = Path.Combine(emailDirectory, "sent");
        var calendarDirectory = Path.Combine(outputDirectory, "calendars");
        Directory.CreateDirectory(inboxDirectory);
        Directory.CreateDirectory(sentDirectory);
        Directory.CreateDirectory(calendarDirectory);

        var userPath = $"users/{Uri.EscapeDataString(emailAddress)}";
        Console.WriteLine($"Target mailbox: {emailAddress}");
        Console.WriteLine($"Period: {cutoff:yyyy-MM-dd HH:mm:ss} UTC to {now:yyyy-MM-dd HH:mm:ss} UTC");
        Console.WriteLine($"Destination: {outputDirectory}");

        var emailCount = await BackupEmailsAsync(
            userPath,
            inboxDirectory,
            sentDirectory,
            cutoff,
            options.UnreadOnly,
            cancellationToken);
        var eventCount = await BackupCalendarAsync(userPath, calendarDirectory, cutoff, now, cancellationToken);

        return new BackupResult(outputDirectory, emailAddress, cutoff, emailCount, eventCount);
    }

    private async Task<int> BackupEmailsAsync(
        string userPath,
        string inboxDirectory,
        string sentDirectory,
        DateTimeOffset cutoff,
        bool unreadOnly,
        CancellationToken cancellationToken)
    {
        var count = 0;
        foreach (var folder in new[]
        {
            (Name: "Inbox", Id: "inbox", DateProperty: "receivedDateTime", Destination: inboxDirectory),
            (Name: "Sent", Id: "sentitems", DateProperty: "sentDateTime", Destination: sentDirectory)
        })
        {
            var filter = $"{folder.DateProperty} ge {cutoff.UtcDateTime:yyyy-MM-ddTHH:mm:ssZ}";
            if (unreadOnly)
            {
                filter += " and isRead eq false";
            }

            var url =
                $"{userPath}/mailFolders/{folder.Id}/messages?$select=id,subject,receivedDateTime,sentDateTime&$filter={Uri.EscapeDataString(filter)}&$top=50";
            var folderCount = 0;
            await foreach (var message in graph.GetCollectionAsync(url, cancellationToken))
            {
                var id = RequiredString(message, "id");
                var subject = GetString(message, "subject") ?? "no-subject";
                var messageDate = GetDateTime(message, folder.DateProperty)
                    ?? throw new InvalidOperationException(
                        $"Message {id} in the {folder.Name} folder does not contain the {folder.DateProperty} date.");
                var filename = $"{messageDate.ToString("yyMMdd-HHmmss", CultureInfo.InvariantCulture)}-{SanitizeFileName(subject)}.eml";
                var destination = GetUniquePath(folder.Destination, filename);
                var temporaryPath = Path.Combine(folder.Destination, $".{Guid.NewGuid():N}.tmp");

                try
                {
                    await graph.DownloadToFileAsync(
                        $"{userPath}/messages/{Uri.EscapeDataString(id)}/$value",
                        temporaryPath,
                        cancellationToken);
                    File.Move(temporaryPath, destination);
                }
                finally
                {
                    if (File.Exists(temporaryPath))
                    {
                        File.Delete(temporaryPath);
                    }
                }

                folderCount++;
                count++;
            }

            Console.WriteLine($"{folder.Name}: {folderCount} email(s) copied");
        }

        return count;
    }

    private async Task<int> BackupCalendarAsync(
        string userPath,
        string destination,
        DateTimeOffset cutoff,
        DateTimeOffset now,
        CancellationToken cancellationToken)
    {
        var start = Uri.EscapeDataString(cutoff.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
        var end = Uri.EscapeDataString(now.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
        var url =
            $"{userPath}/calendarView?startDateTime={start}&endDateTime={end}" +
            "&$select=id,iCalUId,subject,start,end,isAllDay,body,location,organizer,attendees,isCancelled&$top=50";

        var count = 0;
        await foreach (var calendarEvent in graph.GetCollectionAsync(url, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (GetBoolean(calendarEvent, "isCancelled"))
            {
                continue;
            }

            var id = RequiredString(calendarEvent, "id");
            var subject = GetString(calendarEvent, "subject") ?? "no-subject";
            var startDate = GetEventDate(calendarEvent, "start")
                ?? throw new InvalidOperationException($"Event {id} does not contain a start date.");
            var filename = $"{startDate:yyMMdd}-{SanitizeFileName(subject)}.ics";
            var path = GetUniquePath(destination, filename);
            await IcsWriter.WriteEventAsync(path, calendarEvent, cancellationToken);
            count++;
        }

        Console.WriteLine($"Calendar: {count} event(s) exported");
        return count;
    }

    private static string RequiredString(JsonElement element, string property) =>
        GetString(element, property) ??
        throw new InvalidOperationException($"Microsoft Graph response is missing the required property '{property}'.");

    private static string? GetString(JsonElement element, string property) =>
        element.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    private static bool GetBoolean(JsonElement element, string property) =>
        element.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.True;

    private static DateTimeOffset? GetDateTime(JsonElement element, string property)
    {
        var value = GetString(element, property);
        return value is not null && DateTimeOffset.TryParse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal,
            out var dateTime)
                ? dateTime
                : null;
    }

    private static DateOnly? GetEventDate(JsonElement calendarEvent, string property)
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
            return null;
        }

        return DateOnly.FromDateTime(dateTime);
    }

    private static string SanitizeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var safe = new string(value.Select(character =>
            invalid.Contains(character) || char.IsControl(character) ? '_' : character).ToArray())
            .Trim()
            .TrimEnd('.');
        if (safe.Length > 100)
        {
            safe = safe[..100].TrimEnd();
        }

        return string.IsNullOrWhiteSpace(safe) ? "no-subject" : safe;
    }

    private static string GetUniquePath(string directory, string filename)
    {
        var path = Path.Combine(directory, filename);
        if (!File.Exists(path))
        {
            return path;
        }

        var name = Path.GetFileNameWithoutExtension(filename);
        var extension = Path.GetExtension(filename);
        for (var suffix = 2; ; suffix++)
        {
            path = Path.Combine(directory, $"{name}-{suffix}{extension}");
            if (!File.Exists(path))
            {
                return path;
            }
        }
    }
}

internal sealed record BackupResult(
    string OutputDirectory,
    string EmailAddress,
    DateTimeOffset CutoffUtc,
    int EmailCount,
    int EventCount);
