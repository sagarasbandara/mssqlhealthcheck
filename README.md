# MSSQL Daily Health Check

A PowerShell-based SQL Server health check solution that generates a daily HTML report for DBAs, system administrators, and infrastructure teams.

## What It Does

The script automatically identifies whether the target SQL Server is:

- Standalone SQL Server
- SQL Server Failover Cluster node
- Log Shipping enabled server

Based on the detected configuration, it performs a series of health checks and generates a consolidated HTML report.

Checks include:

- Server uptime and instance information
- SQL Server service status
- Database status and availability
- Backup health and recency
- SQL Agent job status
- Disk space availability
- Failed jobs and critical alerts
- Index fragmentation analysis
- High Availability configuration validation
- Log Shipping status (where applicable)

To minimize impact on production environments:

- The tool primarily reads system databases and DMVs
- Large index fragmentation scans can be skipped automatically or manually
- Query timeouts can be configured

## Sample Report

![Health Report Screenshot](docs/report
- Daily DBA health checks
- Operational monitoring
- Infrastructure reviews
- Managed service reporting
- SQL Server environment validation

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
## License

Copyright © 2026 Sagara Bandara

This project is licensed under the MIT License. See the LICENSE file for details.
