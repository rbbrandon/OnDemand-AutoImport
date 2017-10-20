# OnDemand Import Script

This script will read a valid (see below) OnDemand Csv Student Import file, check the local On Demand database for any required additions/modifications, and inserts/updates records as needed. Use this script for automating On Demand imports.

## Getting Started
### Prerequisites:
* SQL Server PowerShell Module (SQLPS) (already installed on OnDemand servers)
* PowerShell v3.0+
* Script Run "As Administrator"
* Script Run Under Account with SQL Database Read/Write Access on the On Demand Server (e.g. "School Admin" account)

### CSV Requirements:
Required CSV headers:
```
student_code,first_name,middle_name,surname,gender,date_of_birth,LBOTE,ATSI,disability_status,EMA,ESL,home_group,year_level
```
Valid CSV Values ([Source](http://webcache.googleusercontent.com/search?q=cache:k5dtnZVSvCAJ:www.vcaa.vic.edu.au/Documents/ondemand/ondteachertrainingguide.pdf+&cd=1&hl=en&ct=clnk&gl=au#31)):

| Field             | Description                            | Accepted Values                                                                     | Mandatory? |
| ----------------: | -------------------------------------- | ------------------------------------------------------------------------------------|------------|
| student_code      | Student code                           | Any string 1-20 characters long.                                                    | Yes        |
| first_name        | Student's first name                   | Any string 1-40 characters long.                                                    | Yes        |
| middle_name       | Student's middle name                  | Any string 0-40 characters long.                                                    | No         |
| surname           | Studnet's surname                      | Any string 1-40 characters long.                                                    | Yes        |
| gender            | Student's gender                       | "M", "MALE", "F", "FEMAL", "FEMALE".                                                | Yes        |
| date_of_birth     | Student's date of birth                | "MMM d yyyy" (e.g. Jan 1 2000), or <br> "d/MM/yyy" (e.g. 1/01/2000).                | Yes        | 
| LBOTE             | Language Background Other Than English | "1", "T", "True", "Y", "Yes" for TRUE; <br> "0", "F", "False", "N", "No" for FALSE. | Yes        |
| ATSI              | Aboriginal or Torres Strait Islander   | "1", "T", "True", "Y", "Yes" for TRUE; <br> "0", "F", "False", "N", "No" for FALSE. | Yes        |
| disability_status | If the student has a disablity or not  | "1", "T", "True", "Y", "Yes" for TRUE; <br> "0", "F", "False", "N", "No" for FALSE. | Yes        |
| EMA               | Education Maintenance Allowance        | "1", "T", "True", "Y", "Yes" for TRUE; <br> "0", "F", "False", "N", "No" for FALSE. | Yes        |
| ESL               | English as a Second Language           | "1", "T", "True", "Y", "Yes" for TRUE; <br> "0", "F", "False", "N", "No" for FALSE. | Yes        |
| home_group        | Students home group                    | Any string 0-40 characters long.                                                    | No         |
| year_level        | Student's year level                   | Numbers 1-12, "F", "UG". <br> (Can be with or without leading zeros).               | Yes        |

### Instructions:
1. Download both `ImportOnDemandUsers.bat` and `ImportOnDemandUsers.ps1`.
2. Modify `ImportOnDemandUsers.bat` to point to a network share that houses your eduHub OnDemand export (`line 9`), and specify the name of the CSV and the log file (`line 12`):
3. Log in to your school's local On Demand server as the user `SchlAdmin`.
4. Copy the script files to the On Demand server (e.g. `C:\Scripts\`).
5. Create a scheduled task to run the `.bat` file ("with highest privileges") whenever you would like the imports to be done. (I do mine at 8am, and repeat every hour for 7 hours)

**Note:** If the script warns you that you need to select a school, please modify the `.bat` file to include the school's ID as per the error message:
e.g If the error says to use school ID of `1234`, then use:

```powershell
PowerShell.exe -ExecutionPolicy Bypass -Command "& '%~dpn0.ps1' 'L:\OnDemand.csv' '1234'" >> L:\OnDemand.log
```

### Examples:

Import "OnDemand.csv" from the "C:\My Files\" directory:

```powershell
.\ImportOnDemandUsers.ps1 "C:\My Files\OnDemand.csv"
```

Import "OnDemand.csv" from the "C:\My Files\" directory, using the SchoolID of "1234" (only use if omitting the SchoolID gives you an error with a list of schools):

```powershell
.\ImportOnDemandUsers.ps1 "C:\My Files\OnDemand.csv" 1234
```

Import "OnDemand.csv" from the "C:\My Files\" directory, and displays additional information:

```powershell
.\ImportOnDemandUsers.ps1 "C:\My Files\OnDemand.csv" -Verbose
```

Use this to run the script as part of a scheduled task:

```powershell
powershell.exe -ExecutionPolicy Bypass -Command "& 'c:\scripts\ImportOnDemandUsers.ps1' 'C:\My Files\OnDemand.csv'"
```

Use this to run the script as part of a scheduled task, and append the output to a log file:

```powershell
powershell.exe -ExecutionPolicy Bypass -Command "& 'c:\scripts\ImportOnDemandUsers.ps1' 'C:\My Files\OnDemand.csv'" >> c:\scripts\ImportOnDemandUsers.log
```

## Notes:
* The script only checks for student code (`STDNT_XID`) to determine a match for existing students.
* I’m not 100% sure of the first "`STDNT_LID`" value, but from looking at my own database, the first record seems to start at "`[SCHL_ID]0000001000`" (e.g. "`13850000001000`"). Why "`1000`" at the end? No idea, but that’s my first record.
* The script sets the deleted indicator (`DLTD_IND`) to `FALSE` for all imported users, as there is no "`ACTV`/`LEFT`" flags in the CSV (which I assume that flag is for). I could set any user not in the CSV to `TRUE`, but I don’t think it matters.
* "`LAST_UPDATED_SQNC`" I assume is just a counter for how many times the record has been updated, starting at "1" for the initial add, and incrementing with each update to the record.
* It just uses an ID of "0" for "`CREATED_USER_ID`"/"`LAST_UPDATED_USER_ID`", as that’s what my records have.

## Authors:

* **Robert Brandon** - *Initial work*