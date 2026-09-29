# M365Backup

A .NET 10 console application that backs up the email and calendar of a single Microsoft 365 mailbox using Microsoft Graph.

## Set up the Entra application

The application uses **application** authentication (client credentials). Configure the Microsoft Graph application permissions `Mail.Read` and `Calendars.Read`, and have an administrator grant consent. This backup does not use `User.Read`. An Exchange application access policy can restrict which mailbox is accessible, if your organization has configured one.

## Configuration

Copy `src/appsettings-example.json` to `src/appsettings.json`, then fill in the credentials and backup settings:

```json
{
  "TenantId": "<Tenant ID>",
  "ClientId": "<Application (client) ID>",
  "ClientSecret": "<client secret value>",
  "Days": 30,
  "OutputDirectory": "C:/_Temp/M365",
  "OutputOverwrite": false,
  "EmailAddress": "<mailbox address to back up>",
  "EmailFolders": {
    "Inbox": {
      "FolderId": "inbox",
      "DateProperty": "receivedDateTime"
    },
    "SentItems": {
      "FolderId": "sentitems",
      "DateProperty": "sentDateTime"
    }
  },
  "EmailReadOnly": false
}
```

`appsettings.json` is ignored by Git. The `Days`, `OutputDirectory`, `OutputOverwrite`, `EmailAddress`, `EmailFolders`, and `EmailReadOnly` settings control the backup:

- `Days`: rolling period of 1 to 3650 days, ending at the time of execution.
- `OutputDirectory`: root directory for the output. Relative paths are resolved from the current working directory.
- `OutputOverwrite`: when `true`, replace files with matching names; when `false`, add a numeric suffix to avoid overwriting.
- `EmailAddress`: mailbox to back up.
- `EmailFolders`: mail folders to export. Each entry name becomes a lowercase subdirectory under `email/`; `FolderId` is the Microsoft Graph mail-folder ID, and `DateProperty` is the message date field used for both filtering and filenames.
- `EmailReadOnly`: when `true`, export only unread emails. This setting does not apply to the calendar.

## Usage

From the repository root:

```powershell
dotnet run --project src -- --config src/appsettings.json
```

The only command-line option is `--config`, which specifies the path to the JSON configuration file.

Emails are exported under `email/` in MIME `.eml` format, using the filename pattern `yyMMdd-HHmmss-Subject.eml`. Non-cancelled calendar events within the configured period are exported individually to `calendars/` in `.ics` format, using the filename pattern `yyMMdd-Subject.ics`.
