namespace M365Backup;

internal sealed record CliOptions(string ConfigPath)
{
    public static CliOptions Parse(string[] args)
    {
        if (args.Length != 2 || args[0] != "--config" || string.IsNullOrWhiteSpace(args[1]))
        {
            throw new ArgumentException("Usage: M365Backup --config <file.json>");
        }

        return new CliOptions(args[1]);
    }
}
