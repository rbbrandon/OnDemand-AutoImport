<#
.SYNOPSIS
This script reads a valid OnDemand Csv Student Import file, and imports/updates users if necessary.

.DESCRIPTION
This script will read a valid (see below) OnDemand Csv Student Import file, check the local On Demand database for any required additions/modifications,
and inserts/updates records as needed. Use this script for automating On Demand imports.

Required CSV headers:
student_code,first_name,middle_name,surname,gender,date_of_birth,LBOTE,ATSI,disability_status,EMA,ESL,home_group,year_level

.NOTES
Author      : Robert Brandon
Created     : 10/10/2017
Last Edited : 18/10/2017
Requires    : SQL Server PowerShell Module (SQLPS);
              PowerShell v3.0+;
              Script Run "As Administrator";
              Script Run Under Account with SQL Database Read/Write Access on the On Demand Server.

(To run as a scheduled task, use  examples 4 or 5.)

.LINK
SQL Server PowerShell Module (SQLPS) : https://docs.microsoft.com/en-us/sql/relational-databases/scripting/import-the-sqlps-module

.EXAMPLE
.\ImportOnDemandUsers.ps1 "C:\My Files\OnDemand.csv"
Imports "OnDemand.csv" from the "C:\My Files\" directory.

.EXAMPLE
.\ImportOnDemandUsers.ps1 "C:\My Files\OnDemand.csv" 1234
Imports "OnDemand.csv" from the "C:\My Files\" directory, and sets the SchoolID to "1234" (only use if omitting the SchoolID gives you an error with a list of schools).

.EXAMPLE
.\ImportOnDemandUsers.ps1 "C:\My Files\OnDemand.csv" -Verbose
Imports "OnDemand.csv" from the "C:\My Files\" directory, and displays additional information.

.EXAMPLE
powershell.exe -ExecutionPolicy Bypass -Command "& 'c:\scripts\ImportOnDemandUsers.ps1' 'C:\My Files\OnDemand.csv'"
Use this to run the script as part of a scheduled task.

.EXAMPLE
powershell.exe -ExecutionPolicy Bypass -Command "& 'c:\scripts\ImportOnDemandUsers.ps1' 'C:\My Files\OnDemand.csv'" >> c:\scripts\ImportOnDemandUsers.log
Use this to run the script as part of a scheduled task, and append the output to a log file.

.PARAMETER CsvFile
Relative or full path to the CSV file to import.

.PARAMETER SchoolID
(Optional) SchoolID of the school to import data for (Note: must be On Demand's SchoolID, not DET's School ID). Only use if omitting the SchoolID gives you an error with a list of schools.

#>

#Requires -RunAsAdministrator
#Requires -Version 3.0

# Import verbosity settings
[CmdletBinding()]
Param(
    [Parameter( Mandatory=$true,
                Position=0,
                ParameterSetName="CsvFile",
                HelpMessage='Please enter the full path to the CSV file to import.')]
    [string] $CsvFile,
    [Parameter( Position=1,
                ParameterSetName="SchoolID",
                HelpMessage='Please specify the On Demand SchoolID.')]
    [string] $SchoolID
)

# Define main script function.
Function Main ($CsvFile, $SchoolID)
{
    Write-Color "`r`n******************************"
    Write-Color "    On Demand CSV Importer    "
    Write-Color "******************************"
    Write-Color "(Runtime: $((Get-Date).ToString("dd/MM/yyyy HH:mm:ss")))`r`n"

    #region Import all required modules
    Write-Output " Importing modules"
    Push-Location
    try {
        Import-Module sqlps -DisableNameChecking
        Write-Green " - SQLPS Loaded."
    }
    catch {
        Write-Red " ERROR: Failed to load SQLPS module"
        Write-Red "        $($_.Exception.Message)"
        return
    }
    Pop-Location
    #endregion

    #region Get CSV information
    Write-Color " Reading OnDemand.csv..."
    try {
        $ImportedCsv = Import-Csv $CsvFile

        Write-Verbose "  $($ImportedCsv | Measure-Object | Select-Object -expand count) Records Read."
        Write-Green " - Done."
    } catch {
        Write-Red "  ERROR: Failed to read $CsvFile"
        Write-Red "         $($_.Exception.Message)"
        return
    }
    #endregion

    #region Get required info from DB
    Write-Color " Retrieving Required ""On Demand"" Information from the local database..."

    #region Get School Information
    try {
        Write-Verbose "  Retrieving list of schools."

        $Schools = Invoke-SqlCmd -Hostname localhost -Database AIM -Query "SELECT SCHL_ID AS SCHOOL_ID, SCHL_NAME AS NAME, LCTN_ADRS_SBRB AS LOCATION FROM SCHOOL"
        if ($null -eq $Schools) {
            Write-Red "  ERROR: Could not obtain School information from the local database."
            Write-Red "         Script must be run on the On Demand server under an account that has write access to the SQL database."
            return
        }

        if ([string]::IsNullOrEmpty($SchoolID)) {
            # No SchoolID was supplied. Try getting it automatically (only if there is 1 school in the DB).
            if ($Schools -is [System.Data.DataRow]) {
                $SchoolID = $Schools.SCHOOL_ID
            } else {
                Write-Red "  ERROR: Could not obtain the SchoolID. More than 1 school was found. please specify the SchoolID."
                Write-Red "         Available Schools:  $(($SchoolID | Format-Table -AutoSize | Out-String).TrimEnd())"
                return
            }
        } elseif ($Schools.SCHOOL_ID -notcontains $SchoolID) {
            # Supplied SchoolID is not in the database.
            Write-Red "  ERROR: Invalid SchoolID."
            Write-Red "         Available Schools:  $(($Schools | Format-Table -AutoSize | Out-String).TrimEnd())"
            return
        } else {
            # Supplied SchoolID is valid.
            Write-Verbose "  Found school with ID ""$SchoolID"" in the local database."
        }
    } catch {
        Write-Red "  ERROR: Could not obtain the SchoolID."
        Write-Red "         Script must be run on the On Demand server under an account that has write access to the SQL database."
        return
    }
    #endregion

    #region Get Year Level Information
    Write-Verbose "  Retrieving list of YearLevels."
    $YearLevels = Invoke-SqlCmd -Hostname localhost -Database AIM -Query "SELECT YEAR_LVL_ID, YEAR_LVL_DSCRPTN FROM YEAR_LEVEL"
    
    if ($null -eq $YearLevels) {
        Write-Color "  ERROR: Could not obtain Year Level information from the local database."
        Write-Color "         Script must be run on the On Demand server under an account that has write access to the SQL database."
        return
    }
    #endregion

    #region Get Existing Student Information
    $Students = Invoke-SqlCmd -Hostname localhost -Database AIM -Query "SELECT * FROM STUDENT"

    # Check if any students were found.
    if ($null -eq $Students) {
        # No students found, must be a new DB. Set initial studentID.
        $Students = @()
        $NextStudentID = "$($SchoolID)0000001000"
    } else {
        $NextStudentID = ($Students | Measure-Object -Property STDNT_LID -Maximum  | select -ExpandProperty Maximum) + 1
    }
    #endregion

    Write-Green " - Done."
    #endregion

    #region Process CSV Records
    Write-Color " Processing CSV Records..."
    $Updates = 0
    $Inserts = 0
    $Errors  = 0

    foreach ($Record in $ImportedCsv) {
        #region Validate record values.
        $Invalid = CheckRecordFail $Record
        if ($Invalid -ne $false) {
            Write-Red $Invalid
            $Errors++
            Continue
        }
        #endregion

        #region "Clean" record values for SQL.
        $Record | Foreach-Object {
            $_.first_name = $_.first_name.Replace("'", "").Replace("""", "");
            $_.middle_name = $_.middle_name.Replace("'", "").Replace("""", "");
            $_.surname = $_.surname.Replace("'", "").Replace("""", "");
            $_.student_code = $_.student_code.Replace("'", "").Replace("""", "");
            $_.home_group = $_.home_group.Replace("'", "").Replace("""", "");
            $_.year_level = $_.year_level.Replace("'", "").Replace("""", "");
        }
        if ($null -eq $Record.home_group) { $Record.home_group = "" }
        if ($null -eq $Record.middle_name) { $Record.middle_name = "" }

        $Birth_Date = [datetime]::ParseExact($Record.date_of_birth, "d/MM/yyyy", $null).ToString("yyyyMMdd HH:mm:ss")

        if ($YearLevels.YEAR_LVL_DSCRPTN -contains $Record.year_level) {
            $Year_Lvl_Id = $YearLevels | Where-Object { $_.YEAR_LVL_DSCRPTN -eq $Record.year_level }  | select -ExpandProperty YEAR_LVL_ID
        } else {
            Write-Red """$($Record.year_level)"" is not a valid year level."
            Continue
        }
    
        $Now = (Get-Date).ToString("yyyyMMdd HH:mm:ss")
        #endregion
        
        #region Perform required database edits.
        Write-Verbose "Processing: $(($Record | Format-Table -AutoSize | Out-String).TrimEnd())"

        # Check if the student code for the current record already exists in the database.
        if ($Students.STDNT_XID -contains $Record.student_code) {
            # Student code already exists - Check if student requires updating.
            $ExistingStudent = $Students | Where-Object { $_.STDNT_XID -eq $Record.student_code }
        
            if ($ExistingStudent.YEAR_LVL_ID -ne $Year_Lvl_Id -or
                $ExistingStudent.STDNT_FRST_NAME -ne $Record.first_name -or
                $ExistingStudent.STDNT_MDL_NAME -ne $Record.middle_name -or
                $ExistingStudent.STDNT_SRNM -ne $Record.surname -or
                $ExistingStudent.GNDR_CD -ne $Record.gender -or
                $ExistingStudent.BIRTH_DT.ToString("yyyyMMdd HH:mm:ss") -ne $Birth_Date -or
                $ExistingStudent.LBOTE_IND -ne [Int32]$Record.LBOTE -or
                $ExistingStudent.ATSI_IND -ne [Int32]$Record.ATSI -or
                $ExistingStudent.DSBLTY_IND -ne [Int32]$Record.disability_status -or
                $ExistingStudent.EMA_IND -ne [Int32]$Record.EMA -or
                $ExistingStudent.ESL_IND -ne [Int32]$Record.ESL -or
                $ExistingStudent.HOME_GRP_NAME -ne $Record.home_group)
            {
                # Student information from CSV is different, therefore student requires updating.
                Write-Color "  Updating $($Record.first_name) $($Record.surname) ($($Record.student_code))'s Student Information..."

                $SQLQuery = "UPDATE STUDENT "`
                          + "SET YEAR_LVL_ID = $Year_Lvl_Id, "`
                              + "STDNT_FRST_NAME = '$($Record.first_name)', "`
                              + "STDNT_MDL_NAME = '$($Record.middle_name)', "`
                              + "STDNT_SRNM = '$($Record.surname)', "`
                              + "GNDR_CD = '$($Record.gender)', "`
                              + "BIRTH_DT = '$($Birth_Date)', "`
                              + "LBOTE_IND = $($Record.LBOTE), "`
                              + "ATSI_IND = $($Record.ATSI), "`
                              + "DSBLTY_IND = $($Record.disability_status), "`
                              + "EMA_IND = $($Record.EMA), "`
                              + "ESL_IND = $($Record.ESL), "`
                              + "HOME_GRP_NAME = '$($Record.home_group)', "`
                              + "LAST_UPDATED_USER_ID = 0, "`
                              + "LAST_UPDATED_DT_TM = '$Now', "`
                              + "LAST_UPDATED_SQNC = $($ExistingStudent.LAST_UPDATED_SQNC + 1) "`
                          + "WHERE STDNT_LID = $($ExistingStudent.STDNT_LID)"

                Write-Verbose "Executing: $SQLQuery"

                try {
                    Invoke-SqlCmd -Hostname localhost -Database AIM -Query $SQLQuery
                    $Updates += 1
                    Write-Green "  - Done."
                } catch {
                    Write-Red "ERROR: Command failed: $SQLQuery"
                }
            } else {
                # Student is in the database, and their details match the CSV. Do nothing.
                Write-Verbose "  $($Record.first_name) $($Record.surname) ($($Record.student_code))'s Student Information is Unchanged."
            }
        } else {
            # Student doesn't exist in the database, therefore must be a new student. Add the student to the database.
            Write-Color "  Adding $($Record.first_name) $($Record.surname) ($($Record.student_code))'s Student Information..."

            # Prepare the SQL Query to run.
            $SQLQuery = "INSERT INTO STUDENT (STDNT_LID, YEAR_LVL_ID, SCHL_ID, STDNT_XID, STDNT_EXTRNL_XID, STDNT_FRST_NAME, STDNT_MDL_NAME, "`
                      + "STDNT_SRNM, GNDR_CD, BIRTH_DT, LBOTE_IND, ATSI_IND, DSBLTY_IND, EMA_IND, ESL_IND, HOME_GRP_NAME, CREATED_USER_ID, "`
                      + "CREATED_DT_TM, LAST_UPDATED_USER_ID, LAST_UPDATED_DT_TM, LAST_UPDATED_SQNC, DLTD_IND) "`
                      + "VALUES ($NextStudentID, $Year_Lvl_Id, $SchoolID, '$($Record.student_code)', '$($Record.student_code)', "`
                      + "'$($Record.first_name)', '$($Record.middle_name)', '$($Record.surname)', '$($Record.gender)', '$Birth_Date', "`
                      + "$($Record.LBOTE), $($Record.ATSI), $($Record.disability_status), $($Record.EMA), $($Record.ESL), '$($Record.home_group)', "`
                      + "0, '$Now', 0, '$Now', 1, 0)"

            Write-Verbose "Executing: $SQLQuery"

            try {
                Invoke-SqlCmd -Hostname localhost -Database AIM -Query $SQLQuery
                $Inserts += 1
                $NextStudentID += 1
                Write-Green "  - Done."
            } catch {
                Write-Red "ERROR: Command failed: $SQLQuery"
            }
        }
        #endregion
    }

    Write-Green " - Done."
    #endregion

    #region Display Results
    Write-Output "`r`n$($ImportedCsv.Count) total records checked:"
    if ($Updates -gt 0) {
        Write-Green "- $Updates Existing Student$(if ($Updates -ne 1) { "s" }) Updated."
    } else {
        Write-Color "- 0 Existing students updated."
    }

    if ($Inserts -gt 0) {
        Write-Green "- $Inserts New Student$(if ($Inserts -ne 1) { "s" }) Added."
    } else {
        Write-Color "- 0 New students added."
    }


    if ($Errors -gt 0) {
        Write-Red "- $Errors Record$(if ($Errors -ne 1) { "s" }) ignored due to errors. :("
    } else {
        Write-Green "- 0 Records with errors. :)"
    }
    #endregion
}

#region Required Custom Functions.
Function EndScript {
    #if (-Not $psISE) {
	#    Write-Host " Press any key to exit ..."
	#    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	#    cls
    #}
	exit
}

Function CheckRecordFail ($Record) {
    $ValidBool        = "0", "1"
    $ValidGenders     = "MALE", "FEMAL"
    $ValidYearLevels  = "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "F", "UG"
    $ValidationFailed = $false
    $ErrorMsg         = ""

    if ([string]::IsNullOrEmpty($Record.ATSI) -or $ValidBool -notcontains $Record.ATSI) {
        $ErrorMsg += "`r`nERROR: ""ATSI"" (Aboriginal or Torres Strait Islander) flag must be either a 0, or 1."
        $ValidationFailed = $true
    }
    if ([string]::IsNullOrEmpty($Record.disability_status) -or $ValidBool -notcontains $Record.disability_status) {
        $ErrorMsg += "`r`nERROR: ""disability_status"" flag must be either a 0, or 1."
        $ValidationFailed = $true
    }
    if ([string]::IsNullOrEmpty($Record.EMA) -or $ValidBool -notcontains $Record.EMA) {
        $ErrorMsg += "`r`nERROR: ""EMA"" (Education Maintenance Allowance) flag must be either a 0, or 1."
        $ValidationFailed = $true
    }
    if ([string]::IsNullOrEmpty($Record.ESL) -or $ValidBool -notcontains $Record.ESL) {
        $ErrorMsg += "`r`nERROR: ""ESL"" (English as a Second Language) flag must be either a 0, or 1."
        $ValidationFailed = $true
    }
    if ([string]::IsNullOrEmpty($Record.LBOTE) -or $ValidBool -notcontains $Record.LBOTE) {
        $ErrorMsg += "`r`nERROR: ""LBOTE"" (Language Background Other Than English) flag must be either a 0, or 1."
        $ValidationFailed = $true
    }
    if ([string]::IsNullOrEmpty($Record.student_code) -or $Record.student_code.Length -gt 20) {
        $ErrorMsg += "`r`nERROR: ""student_code"" must be 1-20 characters in length."
        $ValidationFailed = $true
    }
    if ([string]::IsNullOrEmpty($Record.first_name) -or $Record.first_name.Length -gt 40) {
        $ErrorMsg += "`r`nERROR: ""first_name"" must be 1-40 characters in length."
        $ValidationFailed = $true
    }
    if ([string]::IsNullOrEmpty($Record.surname) -or $Record.surname.Length -gt 40) {
        $ErrorMsg += "`r`nERROR: ""surname"" must be 1-40 characters in length."
        $ValidationFailed = $true
    }
    if ($Record.middle_name.Length -gt 40) {
        $ErrorMsg += "`r`nERROR: ""middle_name"" must be 0-40 characters in length."
        $ValidationFailed = $true
    }
    if ($Record.home_group.Length -gt 40) {
        $ErrorMsg += "`r`nERROR: ""home_group"" must be 0-40 characters in length."
        $ValidationFailed = $true
    }
    if ([string]::IsNullOrEmpty($Record.gender) -or $ValidGenders -notcontains $Record.gender) {
        $ErrorMsg += "`r`nERROR: ""gender"" must be one of the following values: ""$($ValidGenders -join '", "')""."
        $ValidationFailed = $true
    }
    if ([string]::IsNullOrEmpty($Record.year_level) -or $ValidYearLevels -notcontains $Record.year_level) {
        $ErrorMsg += "`r`nERROR: ""year_level"" must be one of the following values: ""$($ValidYearLevels -join '", "')""."
        $ValidationFailed = $true
    }
    
    try {
        $Birth_Date = [datetime]::ParseExact($Record.date_of_birth, "d/MM/yyyy", $null)
    } catch {
        $ErrorMsg += "`r`nERROR: date_of_birth ""$($Record.date_of_birth)"" is not a valid date. Must be in the format ""d/MM/yyyy""."
        $ValidationFailed = $true
    }

    if ($ValidationFailed) {
        $ErrorMsg = "ERROR: The Following CSV Record is Invalid: $(($Record | Format-Table -AutoSize | Out-String).TrimEnd())" + $ErrorMsg
        return $ErrorMsg
    } else {
        return $false
    }
}

function Write-Color([String[]]$Text, [ConsoleColor[]]$Color = @()) {
    if ($Color.Count -ge $Text.Count) {
        for ($i = 0; $i -lt $Text.Length; $i++) {
            Write-Host $Text[$i] -ForegroundColor $Color[$i] -NoNewLine
        } 
    } else {
        for ($i = 0; $i -lt $Color.Length ; $i++) {
            Write-Host $Text[$i] -ForegroundColor $Color[$i] -NoNewLine
        }
        for ($i = $Color.Length; $i -lt $Text.Length; $i++) {
            Write-Host $Text[$i] -NoNewLine
        }
    }

    # Write-Host ends its lines with "`n" (unix/shell format). So by writing "`r" first it'll end the line with "`r`n", which is Windows format.
    Write-Host "`r"
}

Function Write-Red ($Message) {
    Write-Color $Message -Color "Red"
}

Function Write-Green ($Message) {
    Write-Color $Message -Color "DarkGreen"
}
#endregion

# Execute Script
Main $CsvFile $SchoolID

EndScript