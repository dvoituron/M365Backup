using System.Text.Json;

namespace M365Backup;

internal sealed record AppCredentials(
    string TenantId,
    string ClientId,
    string ClientSecret,
    int Days,
    string OutputDirectory,
    bool OutputOverwrite,
    string EmailAddress,
    IReadOnlyDictionary<string, EmailFolderConfiguration> EmailFolders,
    bool EmailReadOnly)
{
    public static AppCredentials Load(string configPath)
    {
        var fullPath = Path.GetFullPath(configPath);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Configuration file not found: {fullPath}", fullPath);
        }

        using var stream = File.OpenRead(fullPath);
        var configuration = JsonSerializer.Deserialize<ConfigurationFile>(
            stream,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        if (configuration is null)
        {
            throw new InvalidOperationException("The configuration file is empty or invalid.");
        }

        var tenantId = Required(configuration.TenantId, "TenantId");
        var clientId = Required(configuration.ClientId, "ClientId");
        var clientSecret = Required(configuration.ClientSecret, "ClientSecret");
        var emailAddress = Required(configuration.EmailAddress, "EmailAddress");
        var outputDirectory = Required(configuration.OutputDirectory, "OutputDirectory");

        if (IsPlaceholder(tenantId) ||
            IsPlaceholder(clientId) ||
            IsPlaceholder(clientSecret) ||
            IsPlaceholder(emailAddress))
        {
            throw new InvalidOperationException("Replace the configuration placeholders with actual values.");
        }

        if (configuration.Days is not (>= 1 and <= 3650))
        {
            throw new InvalidOperationException("Days must be a number between 1 and 3650.");
        }

        if (configuration.OutputOverwrite is null)
        {
            throw new InvalidOperationException("The OutputOverwrite setting is required.");
        }

        if (configuration.EmailReadOnly is null)
        {
            throw new InvalidOperationException("The EmailReadOnly setting is required.");
        }

        if (configuration.EmailFolders is null || configuration.EmailFolders.Count == 0)
        {
            throw new InvalidOperationException("EmailFolders must contain at least one folder.");
        }

        var emailFolders = new Dictionary<string, EmailFolderConfiguration>(StringComparer.OrdinalIgnoreCase);
        foreach (var (name, folder) in configuration.EmailFolders)
        {
            if (string.IsNullOrWhiteSpace(name) ||
                name is "." or ".." ||
                name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                name.Contains('/') ||
                name.Contains('\\'))
            {
                throw new InvalidOperationException($"Invalid EmailFolders entry name: '{name}'.");
            }

            if (folder is null)
            {
                throw new InvalidOperationException($"EmailFolders entry '{name}' is missing its settings.");
            }

            var folderId = Required(folder.FolderId, $"EmailFolders.{name}.FolderId");
            var dateProperty = Required(folder.DateProperty, $"EmailFolders.{name}.DateProperty");
            if (!IsGraphPropertyName(dateProperty))
            {
                throw new InvalidOperationException(
                    $"EmailFolders.{name}.DateProperty must be a valid Microsoft Graph property name.");
            }

            if (!emailFolders.TryAdd(name, new EmailFolderConfiguration(folderId, dateProperty)))
            {
                throw new InvalidOperationException($"Duplicate EmailFolders entry name: '{name}'.");
            }
        }

        return new AppCredentials(
            tenantId,
            clientId,
            clientSecret,
            configuration.Days!.Value,
            outputDirectory,
            configuration.OutputOverwrite.Value,
            emailAddress,
            emailFolders,
            configuration.EmailReadOnly.Value);
    }

    private static string Required(string? value, string name) =>
        !string.IsNullOrWhiteSpace(value)
            ? value
            : throw new InvalidOperationException($"The {name} setting is required.");

    private static bool IsPlaceholder(string value) =>
        value.StartsWith("YOUR_", StringComparison.OrdinalIgnoreCase) ||
        (value.StartsWith('<') && value.EndsWith('>'));

    private static bool IsGraphPropertyName(string value) =>
        (char.IsAsciiLetter(value[0]) || value[0] == '_') &&
        value.All(character => char.IsAsciiLetterOrDigit(character) || character == '_');

    private sealed class ConfigurationFile
    {
        public string? TenantId { get; init; }
        public string? ClientId { get; init; }
        public string? ClientSecret { get; init; }
        public int? Days { get; init; }
        public string? OutputDirectory { get; init; }
        public bool? OutputOverwrite { get; init; }
        public string? EmailAddress { get; init; }
        public Dictionary<string, EmailFolderFile?>? EmailFolders { get; init; }
        public bool? EmailReadOnly { get; init; }
    }

    private sealed class EmailFolderFile
    {
        public string? FolderId { get; init; }
        public string? DateProperty { get; init; }
    }
}

internal sealed record EmailFolderConfiguration(string FolderId, string DateProperty);
