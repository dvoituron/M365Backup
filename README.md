# M365Backup

A .NET 10 console application that backs up the email and calendar of a single Microsoft 365 mailbox using Microsoft Graph.

## Set up the Entra application

The application uses **application** authentication (client credentials). Configure the Microsoft Graph application permissions `Mail.Read` and `Calendars.Read`, and have an administrator grant consent. This backup does not use `User.Read`. An Exchange application access policy can restrict which mailbox is accessible, if your organization has configured one.

## Configuration

Copy `appsettings.example.json` to `appsettings.json`, then fill in:

```json
{
  "TenantId": "<Tenant ID>",
  "ClientId": "<Application (client) ID>",
  "ClientSecret": "<client secret value>",
  "EmailAddress": "<mailbox address to back up>"
}
```

`appsettings.json` is ignored by Git. You can also pass the path to another file with `--config`.

## Usage

From this directory:

```powershell
dotnet run -- --config appsettings.json --days 30 --output C:\Backup\M365 --unread-only false
```

Options:

- `--config`: path to the configuration JSON file.
- `--days`: rolling period of 1 to 3650 days, ending at the time of execution.
- `--output`: output root directory.
- `--unread-only`: `true` to copy only unread emails; defaults to `false`. This filter does not apply to the calendar.

Emails are exported to `email/inbox` and `email/sent` in MIME `.eml` format, using the filename pattern `yyMMdd-HHmmss-Subject.eml`. Non-cancelled calendar events within the period are exported individually to `calendars` in `.ics` format, using the filename pattern `yyMMdd-Subject.ics`. If multiple items have the same name, a numeric suffix is added to prevent overwriting.
