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

- `Days`: rolling period of 1 to 3650 days, ending at the time of execution. Calendar requests are split into windows of at most 1825 days, as required by Microsoft Graph.
- `OutputDirectory`: root directory for the output. Relative paths are resolved from the current working directory.
- `OutputOverwrite`: when `true`, replace existing `.eml` and `.ics` files with matching names; when `false`, skip exporting an item if its output file already exists.
- `EmailAddress`: mailbox to back up.
- `EmailFolders`: mail folders to export. Each entry is saved under `email/<folder>/<year>/`; folder names use the lowercase entry name, except `SentItems`, which is saved under `email/sent/`. `FolderId` is the Microsoft Graph mail-folder ID, and `DateProperty` is the message date field used for filtering, filenames, and year directories.
- `EmailReadOnly`: when `true`, export only read emails; when `false`, export all emails. This setting does not apply to the calendar.

## Usage

From the repository root:

```powershell
dotnet run --project src -- --config src/appsettings.json
```

The only command-line option is `--config`, which specifies the path to the JSON configuration file.

Emails are exported under `email/<folder>/<year>/` in MIME `.eml` format, using the filename pattern `yyMMdd-HHmmss-Subject.eml` (for example, `email/inbox/2026/` and `email/sent/2026/`). Non-cancelled calendar events within the configured period are exported individually to `calendars/<year>/` in `.ics` format, using the filename pattern `yyMMdd-Subject.ics`.

If Microsoft Graph cannot convert a particular message to MIME (`ErrorMimeContentConversionFailed`), that email is reported as failed and the backup continues with the remaining messages. Failures exporting individual emails or calendar events are logged with the item's date, time, and subject. The run exits with an error status and identifies the incomplete backup; the affected item must be handled separately.

For scheduled execution on a Synology NAS, see the [Synology Task Scheduler guide](SYNOLOGY.md).

## Set up the Entra application: step by step

You need access to the Microsoft Entra admin center and an administrator who can grant tenant-wide consent to application permissions.

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com/) and go to **Identity > Applications > App registrations**.
2. Select **New registration**. Enter a name, such as `M365Backup`.
3. Under **Supported account types**, select **Accounts in this organizational directory only** (single tenant). This application uses client credentials and does not need a redirect URI, so leave **Redirect URI** blank.
4. Select **Register**. On the app's **Overview** page, copy:
   - **Application (client) ID** into `ClientId`.
   - **Directory (tenant) ID** into `TenantId`.
5. Go to **Certificates & secrets > Client secrets > New client secret**. Add a description, choose an expiry that complies with your organization's policy, and select **Add**.
6. Copy the secret's **Value** immediately into `ClientSecret`. The value is shown only once; do not use the **Secret ID**. Store the value securely and never commit or share it. Create a replacement before it expires.
7. Go to **API permissions > Add a permission > Microsoft Graph > Application permissions**. Search for and add `Mail.Read` and `Calendars.Read`. These must be **Application permissions**, not Delegated permissions.
8. Select **Grant admin consent for _your organization_** and confirm. An administrator must do this before the application can access mailbox data. The permissions can allow access to mailboxes across the tenant; restrict the application's mailbox scope according to your organization's security policy before running it.
9. Copy `src/appsettings-example.json` to `src/appsettings.json`. Set `TenantId`, `ClientId`, and `ClientSecret` to the values recorded above, and set `EmailAddress` to the mailbox to back up. Adjust `Days`, `OutputDirectory`, `OutputOverwrite`, `EmailFolders`, and `EmailReadOnly` as needed.
10. Confirm that `src/appsettings.json` is ignored by Git and keep it private. 

For more information about creating an app registration, see Microsoft's [application registration documentation](https://learn.microsoft.com/entra/identity-platform/quickstart-register-app). For details about client secrets, see [add and manage application credentials](https://learn.microsoft.com/entra/identity-platform/how-to-add-credentials).
