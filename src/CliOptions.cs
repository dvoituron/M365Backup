namespace M365Backup;

internal sealed record CliOptions(
    string ConfigPath,
    int Days,
    string OutputDirectory,
    bool UnreadOnly,
    bool ShowHelp)
{
    public static CliOptions Parse(string[] args)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        for (var index = 0; index < args.Length; index++)
        {
            var key = args[index];
            if (key is "-h" or "--help")
            {
                return new CliOptions("", 0, "", false, true);
            }

            if (!key.StartsWith("--", StringComparison.Ordinal) || index + 1 >= args.Length)
            {
                throw new ArgumentException($"Invalid option or missing value: {key}");
            }

            if (!values.TryAdd(key, args[++index]))
            {
                throw new ArgumentException($"Repeated option: {key}");
            }
        }

        foreach (var key in values.Keys)
        {
            if (key is not ("--config" or "--days" or "--output" or "--unread-only"))
            {
                throw new ArgumentException($"Unknown option: {key}");
            }
        }

        var configPath = Required(values, "--config");
        if (!int.TryParse(Required(values, "--days"), out var days) || days < 1 || days > 3650)
        {
            throw new ArgumentException("--days must be a number between 1 and 3650.");
        }

        var outputDirectory = Required(values, "--output");
        var unreadOnly = false;
        if (values.TryGetValue("--unread-only", out var unreadOnlyValue) &&
            !bool.TryParse(unreadOnlyValue, out unreadOnly))
        {
            throw new ArgumentException("--unread-only must be true or false.");
        }

        return new CliOptions(configPath, days, outputDirectory, unreadOnly, false);
    }

    public static void PrintHelp()
    {
        Console.WriteLine(
            """
            Back up the email and calendar of a Microsoft 365 mailbox.

            Usage:
              dotnet run -- --config <file.json> --days <number> --output <directory> [--unread-only true|false]

            Options:
              --config <file>       JSON file containing TenantId, ClientId, ClientSecret, and EmailAddress
              --days <number>       Number of recent days to retrieve (1 to 3650)
              --output <directory>  Destination root directory
              --unread-only <bool>  Unread emails only (default: false)
              -h, --help            Show this help
            """);
    }

    private static string Required(IReadOnlyDictionary<string, string> values, string key)
    {
        if (!values.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException($"The {key} option is required.");
        }

        return value;
    }
}
