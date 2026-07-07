# MSSQL Daily Health Check

Generates an HTML daily health report for a Microsoft SQL Server instance.

The report is produced by `Generate-MSSQLDailyHealthReport.ps1` and rendered with
`mssql-daily-health-template.html`.

If you use the provided batch file, create the SQL credential file first. The
health check expects `sql-prod-credential.xml` to exist before it runs.

## Prerequisites

- Windows with PowerShell.
- Network access to the SQL Server instance.
- A SQL Server login or Windows account with permission to read server,
  database, backup, disk, job, and DMV metadata.
- Permission to create files in the report output folder. The default output
  folder is `C:\SQL Reports`.

## First-Time Setup: Create the SQL Credential File

Run this once from the project folder:

```bat
Create-SQLCredentialFile.bat
```

Enter the SQL sysadmin username and password when prompted. The script creates:

```text
sql-prod-credential.xml
```

The credential file is encrypted for the Windows user and computer that created
it. If you move the project to another machine or run it under another Windows
account, create the credential file again.

## Run With the Batch File

After `sql-prod-credential.xml` has been created, double-click or run:

```bat
Run-MSSQLDailyHealthReport.bat
```

This runs:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Generate-MSSQLDailyHealthReport.ps1" -UseSqlAuthentication
```

The batch file uses SQL authentication and reads the encrypted credential file
from this project folder.

## Run Directly With PowerShell

Open PowerShell in this folder.

Run against the local/default SQL Server instance with Windows authentication:

```powershell
.\Generate-MSSQLDailyHealthReport.ps1
```

Run against a specific SQL Server instance with Windows authentication:

```powershell
.\Generate-MSSQLDailyHealthReport.ps1 -ServerInstance "SERVERNAME\INSTANCE"
```

Run with SQL authentication using `sql-prod-credential.xml`:

```powershell
.\Generate-MSSQLDailyHealthReport.ps1 -ServerInstance "SERVERNAME\INSTANCE" -UseSqlAuthentication
```

Run with SQL authentication by passing credentials directly:

```powershell
.\Generate-MSSQLDailyHealthReport.ps1 `
  -ServerInstance "SERVERNAME\INSTANCE" `
  -UseSqlAuthentication `
  -SqlUsername "sql_user" `
  -SqlPassword "sql_password"
```

Write the report to a custom folder:

```powershell
.\Generate-MSSQLDailyHealthReport.ps1 -OutputPath "D:\SQL Reports"
```

Open the report automatically after it is created:

```powershell
.\Generate-MSSQLDailyHealthReport.ps1 -OpenReport
```

Skip the index fragmentation section if it is too slow on a large server:

```powershell
.\Generate-MSSQLDailyHealthReport.ps1 -SkipIndexFragmentation
```

## Output

Reports are written as HTML files named like:

```text
MSSQL_Daily_Health_yyyyMMdd_HHmmss.html
```

Unless `-OutputPath` is provided, files are created in:

```text
C:\SQL Reports
```

## Useful Parameters

| Parameter | Default | Description |
| --- | --- | --- |
| `-ServerInstance` | `.` | SQL Server target. The script also tries common local aliases if the first target fails. |
| `-OutputPath` | `C:\SQL Reports` | Folder where the HTML report is written. |
| `-QueryTimeoutSeconds` | `15` | Timeout for each SQL query. |
| `-UseSqlAuthentication` | off | Use SQL authentication instead of Windows authentication. |
| `-SqlCredentialPath` | `.\sql-prod-credential.xml` | Encrypted credential file used with SQL authentication. |
| `-SkipIndexFragmentation` | off | Skips index fragmentation checks. |
| `-OpenReport` | off | Opens the generated HTML report after creation. |

## Troubleshooting

- If PowerShell blocks the script, run it through the provided `.bat` file or use
  `-ExecutionPolicy Bypass` for this run.
- If SQL authentication fails, recreate `sql-prod-credential.xml` using
  `Create-SQLCredentialFile.bat`.
- If Windows authentication fails with an SSPI or SPN error, try SQL
  authentication or confirm the SQL Server service/SPN configuration.
- If the report folder cannot be created, choose a writable location with
  `-OutputPath`.
- If the report takes too long, increase `-QueryTimeoutSeconds` or use
  `-SkipIndexFragmentation`.
