#################################################################################################################
# 
# Version 0.5
# Daniel Morais Silva (aka Danicell)
#
# Requirements: Windows PowerShell Module for Active Directory
# 				You will need a integration file that contain the fields describled below.
#				
#				LegacyCompany = Is a company Name.
#				LegacyAssocNum = Is a official registration of employee. A Registry.
#				SAPAssocNum = Is the User Logon Name. Here we use a ID like 50000009
#				REGION = Is an establishment field! Here our companies is separated with some establishment.
#				EmployeeStatus = It's a old parameter, we don't use this anymore.
#				Title = Is the role of employee!
#				CostCenterDescr = Is the department of employee!
#				SupervisorSAPAssocNum = Is the SAPAssocNum (Above) of the employee!
#				SAPStatus = A Parameter that tells to me if I need to disable the user or no. (In case of fired employees)
#
#        \\\....
#        |_    |
#       /(O)---`\
#      /_      -'
#        ]   _,')
#       [_,-'_-'(
#      (_).-'    \
#      |/         .
#
#							ps.: Sorry, some variables are in portuguese yet!
#
##################################################################################################################
Clear-Host
## variables
$SQLServer = "AlwaysOnAG_2014.sa.yazaki.local"
$SQLPort = "1433"
$SQLUser = "ad_dat_integration"
$DBName = "AD_DATASUL_INTEGRATION"
$DBTAB = "INTEGRATION"
$SQLPassword = "datint12#"
$FileLocation = "\\ymtatbrfiles2\FTP.YBL"
$DomainOU = "ou=Office Locations,dc=sa,dc=yazaki,dc=local" #OU where users are stored. I Recommend that you create a test OU with some test users to test first.
$QuarantineOU = "OU=Quarantine,OU=Disabled Users,DC=SA,DC=YAZAKI,DC=LOCAL" #OU where fired employee is be stored. Normally, we will put them off of all groups after some months.
$Sep = "|" #This is the separator of CSV (Integration file)

## date variables
$currentDate = [System.DateTime]::Now
$currentDateUtc = $currentDate.ToUniversalTime()

## Here i try to load Ad powershell module, if it's not possible, the script will stop.
## You can install this in Server Manager.
Try
{
  Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
  Write-Host "[ERROR]`t ActiveDirectory Module não pôde ser carregado. Script stop!"
  Exit 1
}

## Open Database connection
$conn = New-Object System.Data.SqlClient.SqlConnection("data source=$SQLServer,$SQLPort;Initial catalog=$DBName;uid=$SQLUser;pwd=$SQLPassword;")
$conn.Open()

## These part of script is resposible to get the last created file in some folder.
## This is usefull to get the last integration file that our HR system generated.
$lastfile = gci $FileLocation | sort LastWriteTime | select -last 1
$lastfilenew = $lastfile.FullName


import-csv $lastfilenew -Delimiter $Sep | foreach {

## Here i just clear some variables in each run of this for.
Clear-Variable CSQL* -Scope Global
Clear-Variable strOU* -Scope Global
Clear-Variable LegacyComp* -Scope Global
Clear-Variable LegacyAss* -Scope Global
Clear-Variable IDUser* -Scope Global
Clear-Variable Estab* -Scope Global
Clear-Variable EmployeeStat* -Scope Global
Clear-Variable Cargo1* -Scope Global
Clear-Variable Setor1* -Scope Global
Clear-Variable Manager* -Scope Global

## This is a translator of each field in the Integration File.
$LegacyComp = $_.LegacyCompany
$LegacyAss = $_.LegacyAssocNum
$IDUser = $_.SAPAssocNum
$Estab = $_.REGION
$EmployeeStat = $_.EmployeeStatus
$Cargo1 = $_.Title
$Setor1 = $_.CostCenterDescr
$Manager = $_.SupervisorSAPAssocNum
$Situacao = $_.SAPStatus

## Parameters to transform some data in Title Case. :) (Workarounds)
$Cargo2 = $Cargo1.ToLower()
$Setor2 = $Setor1.ToLower()
$Cargo = (([system.globalization.cultureinfo]::CurrentCulture).TextInfo).ToTitleCase($Cargo2)
$Setor = (([system.globalization.cultureinfo]::CurrentCulture).TextInfo).ToTitleCase($Setor2)
If ($Setor -eq "Ti Central"){ $Setor = "TI Central" } #Workaround to UpperCase TI Central, because is in lowercase in our TOTVS. :)

## Let the games begin! :p
## Ok, now we will begin all the changes in users that is listed in HR system integration file.
$user = Get-ADUser -SearchBase $DomainOU -Filter {(sAMAccountName -eq $IDUser) -and (enabled -eq $true)} -Properties employeeType,employeeNumber,employeeID,department,description,title,Manager
$sAMAccountName = $user.sAMAccountName
$employeeType = $user.employeeType
If ($sAMAccountName -eq $IDUser){
	Start-Sleep -m 500 ##Just half second to powershell work. I Don't know why, but without this, powershel seems to be crazy. :p
	#Write-Host $sAMAccountName ##Just a simple debug :o
	
## Changing company field!
	If ($employeeType -ne $LegacyComp){
		Set-ADUser -Identity $IDUser -Replace @{employeeType=$LegacyComp}
		$CSQLCOMP = $LegacyComp
	}
	
## Changing establishment field!
	If ($user.employeeNumber -ne $Estab){
		Set-ADUser -Identity $IDUser -EmployeeNumber $Estab
		$CSQLESTAB = $Estab
	}
	
## This is the official registration of employee!
	If ($user.employeeID -ne $LegacyAss){
		Set-ADUser -Identity $IDUser -EmployeeID $LegacyAss
		$CSQLLASS = $LegacyAss
	}
	
## This is the role of employee!
	If ($user.title -ne $Cargo.ToLower()){
		Set-ADUser -Identity $IDUser -Title $Cargo
		$CSQLCARGO = $Cargo
	}
	
## Department of employee!
	If ($user.department -ne $Setor.ToLower() -and $Setor -notlike "AFASTADO*"){
		Set-ADUser -Identity $IDUser -Department $Setor
		$CSQLSETOR = $Setor
	}
	
## Changing Manager of employee!
	If ($ManagerUser -notlike $NULL){
		$ManagerUser = Get-ADUser -Identity $Manager
		$ManagerDist = $ManagerUser.DistinguishedName
		If ($user.Manager -ne $ManagerDist){
			Set-ADUser -Identity $IDUser -Manager $Manager
			$CSQLMANAG = $Manager
		}
	}
	
## Here is another important part of this script!
## Here we will disable a employee that is no longer part of the company!
	If ($Situacao -eq 'T'){ #A fired employee is identifyed with a capitol T in our integration file.
	   	Disable-ADAccount -Identity $IDUser
		$userd = Get-ADUser -Identity $IDUser
		Set-ADUser $userd -Enabled $false
		$strOU = $userd.DistinguishedName
		$strDataMovido_day = $currentDate.Day
		$strDataMovido_month = $currentDate.Month
		$strDataMovido_year = $currentDate.Year
		$strDataMovido = [string]$strDataMovido_year + "-" + [string]$strDataMovido_month + "-" + [string]$strDataMovido_day
		Set-ADUser $userd -Clear "extensionattribute1"
		Set-ADUser $userd -Add @{extensionAttribute1 = $strDataMovido}
		Set-ADUser $userd -Clear "extensionattribute2"
		Set-ADUser $userd -Add @{extensionAttribute2 = "$strOU"}
		Move-ADObject $userd -TargetPath $QuarantineOU
		Write-Host "Usuário $IDUser desabilitado"
		$CSQLDESAB = "X"
		}
	
	## These part of script just log all changes in a database that you have configured above.
	If ($strOU -or $CSQLCOMP -or $CSQLESTAB -or $CSQLLASS -or $CSQLCARGO -or $CSQLSETOR -or $CSQLMANAG -or $CSQLDESAB){
		$cmd = $conn.CreateCommand()
		$cmd.CommandText ="INSERT $DBTAB ([saMAccountName],[OUOrigem],[Empresa],[Estab],[Matricula],[Cargo],[Setor],[Manager],[Desabilitado],[ArquivoRef]) VALUES (N'$IDUser', N'$strOU', N'$CSQLCOMP', N'$CSQLESTAB', N'$CSQLLASS', N'$CSQLCARGO', N'$CSQLSETOR', N'$CSQLMANAG', N'$CSQLDESAB', N'$lastfile')"
		$cmd.ExecuteNonQuery()
	}
}

}

## Here i just close the connection to database.
$conn.Close()

#By Daniel Silva
#Feel free to use this, just maintain the credits.
