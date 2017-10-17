# OnDemand Import Script

This script will read a valid (see below) OnDemand Csv Student Import file, check the local On Demand database for any required additions/modifications, and inserts/updates records as needed. Use this script for automating On Demand imports.

## Getting Started
### Prerequisites
* SQL Server PowerShell Module (SQLPS)
* PowerShell v3.0+
* Script Run "As Administrator"
* Script Run Under Account with SQL Database Read/Write Access on the On Demand Server (e.g. "School Admin" account)

Required CSV headers:
```
student_code,first_name,middle_name,surname,gender,date_of_birth,LBOTE,ATSI,disability_status,EMA,ESL,home_group,year_level
```

### Running

Import "OnDemand.csv" from the "C:\My Files\" directory.

```
.\ImportOnDemandUsers.ps1 "C:\My Files\OnDemand.csv"
```

Import "OnDemand.csv" from the "C:\My Files\" directory, using the SchoolID of "1234" (only use if omitting the SchoolID gives you an error with a list of schools).

```
.\ImportOnDemandUsers.ps1 "C:\My Files\OnDemand.csv" 1234
```

Imports "OnDemand.csv" from the "C:\My Files\" directory, and displays additional information.

```
.\ImportOnDemandUsers.ps1 "C:\My Files\OnDemand.csv" -Verbose
```

Use this to run the script as part of a scheduled task.

```
powershell.exe -ExecutionPolicy Bypass -Command "& 'c:\scripts\ImportOnDemandUsers.ps1' 'C:\My Files\OnDemand.csv'"
```

Use this to run the script as part of a scheduled task, and append the output to a log file.

```
powershell.exe -ExecutionPolicy Bypass -Command "& 'c:\scripts\ImportOnDemandUsers.ps1' 'C:\My Files\OnDemand.csv'" >> c:\scripts\ImportOnDemandUsers.log
```


## Authors

* **Robert Brandon** - *Initial work*