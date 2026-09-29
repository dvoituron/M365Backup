using System.Text.Json;

namespace M365Backup;

internal sealed record AppCredentials(string TenantId, string ClientId, string ClientSecret, string EmailAddress)
{
    public static AppCredentials Load(string configPath)
    {
        var fullPath = Path.GetFullPath(configPath);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Configuration file not found: {fullPath}", fullPath);
        }

        using var stream = File.OpenRead(fullPath);
        var credentials = JsonSerializer.Deserialize<AppCredentials>(
            stream,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        if (credentials is null ||
            string.IsNullOrWhiteSpace(credentials.TenantId) ||
            string.IsNullOrWhiteSpace(credentials.ClientId) ||
            string.IsNullOrWhiteSpace(credentials.ClientSecret) ||
            string.IsNullOrWhiteSpace(credentials.EmailAddress) ||
            credentials.TenantId.StartsWith("YOUR_", StringComparison.OrdinalIgnoreCase) ||
            credentials.ClientId.StartsWith("YOUR_", StringComparison.OrdinalIgnoreCase) ||
            credentials.ClientSecret.StartsWith("YOUR_", StringComparison.OrdinalIgnoreCase) ||
            credentials.EmailAddress.StartsWith("YOUR_", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "The configuration file must contain the actual values for TenantId, ClientId, ClientSecret, and EmailAddress. " +
                "ClientSecret refers to the client secret value, not its Secret ID.");
        }

        return credentials;
    }
}
