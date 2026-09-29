# Run M365Backup on a Synology NAS

This guide explains how to schedule backups using the NAS's DSM web interface. The self-contained publish includes the .NET runtime, so you do not need to install .NET on the Synology.

## 1. Understand the NAS platform

DSM is based on Linux. Linux is the operating system; `x86_64` or `aarch64` identifies the processor architecture. The Task Scheduler script will detect both automatically.

## 2. Publish the application for the supported architectures

To let the script choose the correct executable without having to know the processor architecture in advance, publish both architectures **on your computer**, from the repository root:

```powershell
dotnet publish .\src\M365Backup.csproj -c Release -r linux-arm64 --self-contained true -p:PublishSingleFile=true -o .\publish\linux-arm64
dotnet publish .\src\M365Backup.csproj -c Release -r linux-x64   --self-contained true -p:PublishSingleFile=true -o .\publish\linux-x64
```

Each publish contains an executable named `M365Backup`. Both are required to support ARM 64-bit and x86-64 NAS devices automatically; the NAS will run only one of them.

## 3. Create the folders and transfer the files in DSM

1. Open [the DSM portal](https://192.168.1.59:5001/) and sign in.
2. In **File Station**, create the `/volume1/M365Backup` folder (or choose another shared folder that the task account can access), then create the `linux-arm64` and `linux-x64` subfolders.
3. From your computer, upload the `M365Backup` executable:
   - from `publish/linux-arm64` to `/volume1/M365Backup/linux-arm64`;
   - from `publish/linux-x64` to `/volume1/M365Backup/linux-x64`.
4. Also copy your `src/appsettings.json` from your computer to `/volume1/M365Backup/appsettings.json`.
5. Enter your Entra credentials and backup settings. For `OutputDirectory`, use an absolute Linux path such as `/volume1/M365Backup/Data`.
6. In **File Station**, create the destination folders `/volume1/M365Backup/Data` and `/volume1/M365Backup/Logs`.
7. For each `M365Backup` file, open **Properties > Permission** in File Station and allow execution for the DSM account that will run the task. Give this account read access to the executables and `appsettings.json`, and write access to the backup folder. The account must also be able to write to the application folder to create logs.

The `appsettings.json` file contains the client secret. In **Properties > Permission**, restrict access to accounts that need it.

## 4. Create the scheduled task in DSM

1. In DSM, open **Control Panel > Task Scheduler**.
2. Choose **Create > Scheduled Task > User-defined script** (labels may vary slightly depending on your DSM version).
3. Name the task, for example `M365Backup`, then set the desired schedule.
4. Under **Task Settings**, select a DSM account with the permissions described above. Avoid using `root` unless necessary.
5. In the **User-defined script** field, paste the script below. It checks that the operating system is Linux, detects the processor architecture, runs the matching executable, and appends standard output and errors to a dated log:

   ```sh
   #!/bin/sh
   APP_DIR="/volume1/M365Backup"
   LOG_DIR="$APP_DIR/Logs"
   mkdir -p "$LOG_DIR" || exit 1

   if [ "$(uname -s)" != "Linux" ]; then
     echo "Unsupported operating system: $(uname -s) (Linux is required)" >&2
     exit 1
   fi

   case "$(uname -m)" in
     x86_64) APP="$APP_DIR/linux-x64/M365Backup" ;;
     aarch64) APP="$APP_DIR/linux-arm64/M365Backup" ;;
     *)
       echo "Unsupported architecture: $(uname -m)" >&2
       exit 1
       ;;
   esac

   cd "$APP_DIR" || exit 1
   "$APP" --config "$APP_DIR/appsettings.json" >> "$LOG_DIR/backup-$(date +%F).log" 2>&1
   ```

   If File Station did not allow you to enable execution for the files, the task will fail to start. Check their execution permissions.

6. Save the task, select it, and click **Run** to test it. Check the task history in DSM and, in File Station, the log created under `/volume1/M365Backup/Logs/`. A successful backup also includes `Backup completed` in the log.
7. After the test, use File Station to verify that the backup files appear under `/volume1/M365Backup/Data/`.

The NAS must be able to access the internet and Microsoft Graph. Logs are stored in `Logs/backup-YYYY-MM-DD.log`; periodically delete or archive old files using File Station.
