#************** Enter your school details here ******************************************************************

$DriveLetter = "L:\"		#The path (e.g. "L:\" or "D:\OnDemand\") to copy the output too.

#*******************************************************************************************************

# Pick up school number automagically.
$School_Number = [system.environment]::MachineName.Trim().Substring(0,4)

$StudentFile = "D:\eduHub\ST_" + $School_Number + ".csv"
$StudentDeltaFile = "D:\eduHub\ST_" + $School_Number + "_D.csv"
$FamilyFile = "D:\eduHub\DF_" + $School_Number + ".csv"
$FamilyDeltaFile = "D:\eduHub\DF_" + $School_Number + "_D.csv"
$OutPutFile = $DriveLetter + "OnDemand.csv"

# Check if Delta files exist, and if so, include them.
$StudentCSVFiles = @($StudentFile)
if (Test-Path $StudentDeltaFile) { $StudentCSVFiles += $StudentDeltaFile }

$FamilyCSVFiles = @($FamilyFile)
if (Test-Path $FamilyDeltaFile) { $FamilyCSVFiles += $FamilyDeltaFile }

# Import student eduHub data.
$StudentCSV = Import-Csv -Path $StudentCSVFiles | 
	Where-Object { $_.STATUS -eq "ACTV" -or $_.STATUS -eq "FUT" -or $_.STATUS -eq "LVNG" } |
    Foreach-Object {$_.LW_DATE = [DateTime]::ParseExact($_.LW_DATE,"d/MM/yyyy h:mm:ss tt", [System.Globalization.CultureInfo]::InvariantCulture); $_} | 
    Group-Object STKEY | 
    Foreach-Object {$_.Group | Sort-Object LW_DATE | Select-Object -Last 1} |
	select @{ Name = "student_code"; Expression = { $_.STKEY } },
		@{ Name = "first_name"; Expression = { $_.FIRST_NAME } },
		@{ Name = "middle_name"; Expression = { $_.SECOND_NAME } },
		@{ Name = "surname"; Expression = { $_.SURNAME } },
		@{ Name = "gender"; Expression = { $_.GENDER } },
		@{ Name = "date_of_birth"; Expression = { (Get-Date -format d -Date $_.BIRTHDATE) } },
		LOTE_HOME_CODE,
		@{ Name = "LBOTE"; Expression = { "0" } },
		@{ Name = "ATSI"; Expression = { $_.KOORIE } },
		@{ Name = "disability_status"; Expression = { $_.DISABILITY } },
		@{ Name = "EMA"; Expression = { $_.ED_ALLOW } },
		ENG_SPEAK,
		@{ Name = "ESL"; Expression = { "0" } },
		@{ Name = "home_group"; Expression = { $_.HOME_GROUP } },
		FAMILY,
		@{ Name = "year_level"; Expression = { ($_.SCHOOL_YEAR) } } |
	Sort-Object -property student_code

# Import family eduHub data (for LBOTE calculation)
$FamilyCSV = Import-Csv -Path $FamilyCSVFiles | 
    Foreach-Object {$_.LW_DATE = [DateTime]::ParseExact($_.LW_DATE,"d/MM/yyyy h:mm:ss tt", [System.Globalization.CultureInfo]::InvariantCulture); $_} | 
    Group-Object DFKEY | 
    Foreach-Object {$_.Group | Sort-Object LW_DATE | Select-Object -Last 1} |
	select DFKEY,
		LOTE_HOME_CODE_A,
		LOTE_HOME_CODE_B |
	Sort-Object -property DFKEY

# Check family language codes for LBOTE status
foreach ($Student in $StudentCSV) {
    if ($Student.gender -eq "M") {
        $Student.gender = "MALE";
    } elseif ($Student.gender -eq "F") {
        $Student.gender = "FEMAL";
    } else {
        $Student.gender = "";
    }

    # Check family language codes for LBOTE status
    if ($Student.LOTE_HOME_CODE -ne $null -and $Student.LOTE_HOME_CODE -ne "1201") {
        $Student.LBOTE = "1";
    } else {
        if ($Student.family -ne $null) {
            foreach ($Family in $FamilyCSV) {
                if ($Student.FAMILY -eq $Family.DFKEY) {
                    if (($Family.LOTE_HOME_CODE_A -ne $null -and $Family.LOTE_HOME_CODE_A -ne "1201") -or ($Family.LOTE_HOME_CODE_B -ne $null -and $Family.LOTE_HOME_CODE_B -ne "1201")){
                        $Student.LBOTE = "1";
                    } else {
                     $Student.LBOTE = "0";
                    }

                    break;
                }
            }
        } else {
            $Student.LBOTE = "0";
        }
    }

    if($Student.ATSI -eq "K" -or $Student.ATSI -eq "T" -or $Student.ATSI -eq "B" -or $Student.ATSI -eq "U") {
        $Student.ATSI = "1";
    } else {
        $Student.ATSI = "0";
    }

    if($Student.disability_status -eq "Y") {
        $Student.disability_status = "1";
    } else {
        $Student.disability_status = "0";
    }

    if($Student.EMA -eq "Y") {
        $Student.EMA = "1";
    } else {
        $Student.EMA = "0";
    }

    if($Student.ENG_SPEAK -eq "N") {
        $Student.ESL = "1";
    } else {
        $Student.ESL = "0";
    }
}

#Save CSV
$StudentCSV |
    select student_code,
		first_name,
		middle_name,
		surname,
		gender,
		date_of_birth,
		LBOTE,
		ATSI,
		disability_status,
		EMA,
		ESL,
		home_group,
		year_level |
	ConvertTo-Csv -NoTypeInformation |
    % {$_.Replace('"','')} |
    Out-File $OutPutFile