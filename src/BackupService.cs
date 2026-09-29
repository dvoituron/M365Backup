using System.Globalization;
using System.Text.Json;

namespace M365Backup;

internal sealed class BackupService(GraphReadClient graph)
{
    public async Task<BackupResult> RunAsync(
        AppCredentials configuration,
        CancellationToken cancellationToken)
    {
        var now = DateTimeOffset.UtcNow;
        var cutoff = now.AddDays(-configuration.Days);
        var outputDirectory = Path.GetFullPath(configuration.OutputDirectory);
        if (File.Exists(outputDirectory))
        {
            throw new IOException($"The destination path exists but is not a directory: {outputDirectory}");
        }

        var emailDirectory = Path.Combine(outputDirectory, "email");
        var calendarDirectory = Path.Combine(outputDirectory, "calendars");
        Directory.CreateDirectory(emailDirectory);
        Directory.CreateDirectory(calendarDirectory);

        var userPath = $"users/{Uri.EscapeDataString(configuration.EmailAddress)}";
        Console.WriteLine($"Target mailbox: {configuration.EmailAddress}");
        Console.WriteLine($"Period: {cutoff:yyyy-MM-dd HH:mm:ss} UTC to {now:yyyy-MM-dd HH:mm:ss} UTC");
        Console.WriteLine($"Destination: {outputDirectory}");

        var (emailCount, failedEmailCount) = await BackupEmailsAsync(
            userPath,
            emailDirectory,
            configuration.EmailFolders,
            cutoff,
            configuration.EmailReadOnly,
            configuration.OutputOverwrite,
            cancellationToken);
        var (eventCount, failedEventCount) = await BackupCalendarAsync(
            userPath,
            calendarDirectory,
            cutoff,
            now,
            configuration.OutputOverwrite,
            cancellationToken);

        return new BackupResult(
            outputDirectory,
            configuration.EmailAddress,
            cutoff,
            emailCount,
            eventCount,
            failedEmailCount,
            failedEventCount);
    }

    private async Task<(int Count, int FailedCount)> BackupEmailsAsync(
        string userPath,
        string emailDirectory,
        IReadOnlyDictionary<string, EmailFolderConfiguration> folders,
        DateTimeOffset cutoff,
        bool emailReadOnly,
        bool outputOverwrite,
        CancellationToken cancellationToken)
    {
        var count = 0;
        var totalFailedCount = 0;
        foreach (var (name, folder) in folders)
        {
            var folderDirectoryName = name.Equals("SentItems", StringComparison.OrdinalIgnoreCase)
                ? "sent"
                : name.ToLowerInvariant();
            var filter = $"{folder.DateProperty} ge {cutoff.UtcDateTime:yyyy-MM-ddTHH:mm:ssZ}";
            if (emailReadOnly)
            {
                filter += " and isRead eq true";
            }

            var url =
                $"{userPath}/mailFolders/{Uri.EscapeDataString(folder.FolderId)}/messages?$select=id,subject,{folder.DateProperty}&$filter={Uri.EscapeDataString(filter)}&$top=50";
            var folderCount = 0;
            var skippedCount = 0;
            var failedCount = 0;
            await foreach (var message in graph.GetCollectionAsync(url, cancellationToken))
            {
                var id = RequiredString(message, "id");
                var subject = GetString(message, "subject") ?? "no-subject";
                var messageDate = GetDateTime(message, folder.DateProperty)
                    ?? throw new InvalidOperationException(
                        $"Message {id} in the {name} folder does not contain the {folder.DateProperty} date.");
                var filename = $"{messageDate.ToString("yyMMdd-HHmmss", CultureInfo.InvariantCulture)}-{SanitizeFileName(subject)}.eml";
                var destination = Path.Combine(emailDirectory, folderDirectoryName, messageDate.Year.ToString(CultureInfo.InvariantCulture));
                Directory.CreateDirectory(destination);
                var messagePath = GetOutputPath(destination, filename, outputOverwrite);
                if (messagePath is null)
                {
                    skippedCount++;
                    continue;
                }

                var temporaryPath = Path.Combine(destination, $".{Guid.NewGuid():N}.tmp");

                try
                {
                    await graph.DownloadToFileAsync(
                        $"{userPath}/messages/{Uri.EscapeDataString(id)}/$value",
                        temporaryPath,
                        cancellationToken);
                    File.Move(temporaryPath, messagePath, outputOverwrite);
                }
                catch (GraphRequestException exception) when (
                    exception.StatusCode == System.Net.HttpStatusCode.InternalServerError &&
                    exception.GraphErrorCode == "ErrorMimeContentConversionFailed")
                {
                    failedCount++;
                    Console.Error.WriteLine(
                        $"Failed to export {name} email: date/time {messageDate.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss 'UTC'", CultureInfo.InvariantCulture)}, subject '{subject}', ID {id}: {exception.Message}");
                    continue;
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

            Console.WriteLine(
                $"{name}: {folderCount} email(s) copied" +
                (skippedCount > 0 ? $", {skippedCount} existing email(s) skipped" : "") +
                (failedCount > 0 ? $", {failedCount} email(s) failed" : ""));
            totalFailedCount += failedCount;
        }

        return (count, totalFailedCount);
    }

    private async Task<(int Count, int FailedCount)> BackupCalendarAsync(
        string userPath,
        string destination,
        DateTimeOffset cutoff,
        DateTimeOffset now,
        bool outputOverwrite,
        CancellationToken cancellationToken)
    {
        var count = 0;
        var skippedCount = 0;
        var failedCount = 0;
        var eventIds = new HashSet<string>(StringComparer.Ordinal);
        var windowStart = cutoff;
        while (windowStart < now)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var windowEnd = windowStart.AddDays(1825);
            if (windowEnd > now)
            {
                windowEnd = now;
            }

            var start = Uri.EscapeDataString(
                windowStart.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
            var end = Uri.EscapeDataString(
                windowEnd.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
            var url =
                $"{userPath}/calendarView?startDateTime={start}&endDateTime={end}" +
                "&$select=id,iCalUId,subject,start,end,isAllDay,body,location,organizer,attendees,isCancelled&$top=50";

            await foreach (var calendarEvent in graph.GetCollectionAsync(url, cancellationToken))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var id = GetString(calendarEvent, "id") ?? "unknown";
                var subject = GetString(calendarEvent, "subject") ?? "no-subject";
                var eventDateTime = GetEventDateTimeDescription(calendarEvent);
                try
                {
                    id = RequiredString(calendarEvent, "id");
                    if (!eventIds.Add(id) || GetBoolean(calendarEvent, "isCancelled"))
                    {
                        continue;
                    }

                    var startDate = GetEventDate(calendarEvent, "start")
                        ?? throw new InvalidOperationException($"Event {id} does not contain a start date.");
                    var filename = $"{startDate:yyMMdd}-{SanitizeFileName(subject)}.ics";
                    var yearDirectory = Path.Combine(destination, startDate.Year.ToString(CultureInfo.InvariantCulture));
                    Directory.CreateDirectory(yearDirectory);
                    var path = GetOutputPath(yearDirectory, filename, outputOverwrite);
                    if (path is null)
                    {
                        skippedCount++;
                        continue;
                    }

                    await IcsWriter.WriteEventAsync(path, calendarEvent, cancellationToken);
                    count++;
                }
                catch (Exception exception) when (
                    exception is IOException or UnauthorizedAccessException or InvalidOperationException or ArgumentException)
                {
                    failedCount++;
                    Console.Error.WriteLine(
                        $"Failed to export calendar event: date/time {eventDateTime}, subject '{subject}', ID {id}: {exception.Message}");
                }
            }

            windowStart = windowEnd;
        }

        Console.WriteLine(
            $"Calendar: {count} event(s) exported" +
            (skippedCount > 0 ? $", {skippedCount} existing event(s) skipped" : "") +
            (failedCount > 0 ? $", {failedCount} event(s) failed" : ""));
        return (count, failedCount);
    }

    private static string RequiredString(JsonElement element, string property) =>
        GetString(element, property) ??
        throw new InvalidOperationException($"Microsoft Graph response is missing the required property '{property}'.");

    private static string? GetString(JsonElement element, string property) =>
        element.ValueKind == JsonValueKind.Object &&
        element.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    private static bool GetBoolean(JsonElement element, string property) =>
        element.ValueKind == JsonValueKind.Object &&
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

    private static string GetEventDateTimeDescription(JsonElement calendarEvent)
    {
        if (calendarEvent.ValueKind == JsonValueKind.Object &&
            calendarEvent.TryGetProperty("start", out var start) &&
            start.ValueKind == JsonValueKind.Object &&
            start.TryGetProperty("dateTime", out var dateTime) &&
            dateTime.ValueKind == JsonValueKind.String)
        {
            var timeZone = GetString(start, "timeZone");
            return timeZone is null
                ? dateTime.GetString() ?? "unknown"
                : $"{dateTime.GetString()} {timeZone}";
        }

        return "unknown";
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

    private static string? GetOutputPath(string directory, string filename, bool overwrite)
    {
        var path = Path.Combine(directory, filename);
        return overwrite || !File.Exists(path) ? path : null;
    }
}

internal sealed record BackupResult(
    string OutputDirectory,
    string EmailAddress,
    DateTimeOffset CutoffUtc,
    int EmailCount,
    int EventCount,
    int FailedEmailCount,
    int FailedEventCount);
