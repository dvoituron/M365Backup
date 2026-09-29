using Azure.Identity;

namespace M365Backup;

internal static class Program
{
    private static readonly string[] Scopes = ["https://graph.microsoft.com/.default"];

    private static async Task<int> Main(string[] args)
    {
        try
        {
            var options = CliOptions.Parse(args);
            var configuration = AppCredentials.Load(options.ConfigPath);
            var credential = new ClientSecretCredential(
                configuration.TenantId,
                configuration.ClientId,
                configuration.ClientSecret);

            using var cancellation = new CancellationTokenSource();
            Console.CancelKeyPress += (_, eventArgs) =>
            {
                eventArgs.Cancel = true;
                cancellation.Cancel();
            };

            using var httpClient = new HttpClient();
            var graph = new GraphReadClient(httpClient, credential, Scopes);
            var service = new BackupService(graph);
            var result = await service.RunAsync(configuration, cancellation.Token);

            Console.WriteLine();
            Console.WriteLine($"Backup completed: {result.OutputDirectory}");
            Console.WriteLine($"Mailbox: {result.EmailAddress}");
            Console.WriteLine($"Emails EML: {result.EmailCount}");
            Console.WriteLine($"ICS events: {result.EventCount}");
            return 0;
        }
        catch (OperationCanceledException)
        {
            Console.Error.WriteLine("Operation canceled.");
            return 1;
        }
        catch (Exception exception)
        {
            for (Exception? cause = exception; cause is not null; cause = cause.InnerException)
            {
                var label = ReferenceEquals(cause, exception) ? "Error" : "Cause";
                Console.Error.WriteLine($"{label} ({cause.GetType().Name}): {cause.Message}");
            }

            return 1;
        }
    }
}
