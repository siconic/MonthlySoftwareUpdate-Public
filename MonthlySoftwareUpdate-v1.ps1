# -----------------------------------------------------------------------------
# Script: MonthlySoftwareUpdate.ps1
# Author: Earl
# Date: 01/01/2014
# Modified: 07/11/2024  
#
#          
# Version 1.0.0
#
# Description: Checks for patch Tuesday updates and deploys them
# 
# Pre-Requisits: An ADR will need to be created. The settings are highly customizable
#        and can be whatever you want. I suggest you use the following;
#		1. Software Updates - 
#			A. Date Released - Last 4 days
#			B. Required - >=1
#			C. Superseded - No
#			D. Update Classifiction - Critical Updates or Security Updates
#
#	The following ARE REQUIRED;
#		1. Deployment Settings - Deploy and Approve
#		2. Do not run this rule automatically
#		3. Alerts - Generate alerts when rule fails
#		4. Deployment Package - Select A deployment Package
# -----------------------------------------------------------------------------
$ScriptVer = "1.0.0"
$LastEdit = "Earl"

# -----------------------------------------------------------------------------
#	ENVIRONMENT SPECIFIC VARIABLES
# -----------------------------------------------------------------------------
$SiteCode = ""
$SiteServer = ""
$UpdateShare = "" #DFS Share OR Drive letter
$Collection1 = ""
$Collection2 = ""
$Collection3 = ""
$Collection4 = ""
$Collection5 = ""
$Collection6 = ""
$Collection7 = ""
$Collection8 = ""
$ADRName = "ADR for Monthly Security Updates"
$DPGroup = "All Distribution Points"

# -----------------------------------------------------------------------------
#	GLOBAL VARIABLES
# -----------------------------------------------------------------------------
$CurrentYear = (Get-Date).Year
$CurrentMonth = (Get-Culture).DateTimeFormat.GetMonthName((Get-Date).Month)
$UpdateGroupName = "Patch Compliance - {0} {1} Security Updates" -f $CurrentYear, $CurrentMonth
$DistributeContentAuto = $true

# -----------------------------------------------------------------------------
#	Email Settings
# -----------------------------------------------------------------------------

#SMPT Settings
$SMTPServerName = ""
$SMTPEmailAddress = ""

#Primary Distribution list or email address'
$SCCMTeam = @()
$ServerTeam = @()
$Management = @()


# -----------------------------------------------------------------------------
#	Load ConfigManager (may throw an error)
# -----------------------------------------------------------------------------

Set-Location "$env:SMS_ADMIN_UI_PATH\..\"
Import-Module .\ConfigurationManager.psd1
Set-Location $SiteCode':'

# -----------------------------------------------------------------------------
#	Setting the Deployment timeframes
# -----------------------------------------------------------------------------
#USAGE - Add hours form midnight (12:00am) on Patch Tuesday

#
$Collection1StartDays = 0
$Collection1StartHour = 0
$Collection1DeadlineDays = 0
$Collection1DeadlineHour = 0

#
$Collection2StartDays = 0
$Collection2StartHour = 0
$Collection2DeadlineDays = 0
$Collection2DeadlineHour = 0

#
$Collection3StartDays = 0
$Collection3StartHour = 0
$Collection3DeadlineDays = 0
$Collection3DeadlineHour = 0

#
$Collection4StartDays = 0
$Collection4StartHour = 0
$Collection4DeadlineDays = 0
$Collection4DeadlineHour = 0 

#
$Collection5StartDays = 0
$Collection5StartHour = 0
$Collection5DeadlineDays = 0
$Collection5DeadlineHour = 0 

#
$Collection6StartDays = 0
$Collection6StartHour = 0
$Collection6DeadlineDays = 0
$Collection6DeadlineHour = 0 

#
$Collection7StartDays = 0
$Collection7StartHour = 0
$Collection7DeadlineDays = 0
$Collection7DeadlineHour = 0 

#
$Collection8StartDays = 0
$Collection8StartHour = 0
$Collection8DeadlineDays = 0
$Collection8DeadlineHour = 0 

# ----------------------------------------------------------------------------- 
#  Functions
# ----------------------------------------------------------------------------- 

Function Get-AutomaticDeploymentRule 
{
    [CmdletBinding()]
    Param(
         [Parameter(Mandatory=$false)]
         [ValidateNotNullOrEmpty()]
            [string[]]$Name
    )
    $Class = "SMS_AutoDeployment"
    
    Try {
        $filter = $null
        if($Name) {
            $filter += @("Name LIKE '"+ ($Name -join "' OR Name LIKE '") + "'")
        }
                
        $filter = $filter -join " AND "
        Get-WmiObject -Namespace "root\SMS\Site_$SiteCode" -Class $Class -Filter $filter -ErrorAction STOP -ComputerName $SiteServer | ForEach-Object {[wmi]$_.Path | Select-Object * -ExcludeProperty "__*", "Scope", "Options", "ClassPath", "Properties", "Qualifiers", "Site", "Container"}

     }
     Catch {
        Write-Host "Error: $($_.Exception.Message)"
     }

}

Function Start-AutomaticDeploymentRule
{
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact="High")]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
            $ADR
    )
    Process {
        Foreach ($rule in $ADR) {
            if ($pscmdlet.ShouldProcess($rule.Name)) {
                Try {
                    ([wmi]$rule.Path).EvaluateAutoDeployment()
                } Catch {
                    Write-Host "Error: $($_.Exception.Message)"
                }
            }
        }
    }
}



Function Get-SoftwareUpdateGroup
{
    [CmdletBinding()]
    Param(
         [Parameter(Mandatory=$false)]
         [ValidateNotNullOrEmpty()]
            [string[]]$Name
    )
    $Class = "SMS_AuthorizationList"
    
    Try {
        $filter = $null
        if($Name) {
            $filter += @("LocalizedDisplayName LIKE '"+ ($Name -join "' OR LocalizedDisplayName LIKE '") + "'")
        }
                
        $filter = $filter -join " AND "
        Get-WmiObject -Namespace "root\SMS\Site_$SiteCode" -Class $Class -Filter $filter -ErrorAction STOP -ComputerName $SiteServer | ForEach-Object {[wmi]($_).Path | Select-Object * -ExcludeProperty "__*", "Scope", "Options", "ClassPath", "Properties", "Qualifiers", "Site", "Container"}

     }
     Catch {
        Write-Host "Error: $($_.Exception.Message)"
     }

}

Function Rename-SoftwareUpdateGroup
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
            $UpdateGroup,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
            [string]$NewName
    )
      
        if ($UpdateGroup.SystemProperties.Item("__CLASS").Value -eq "SMS_AuthorizationList") {
            $info = $UpdateGroup.LocalizedInformation
            $info[0].DisplayName = $NewName
            Set-WmiInstance -Path $UpdateGroup.Path -Arguments @{LocalizedInformation = $info}
        } else {
            throw "Specified Update Group is not a valid object"
        }
}

Function Get-Collection
{
    [CmdletBinding()]
    Param(
         [Parameter(Mandatory=$false,
            HelpMessage="Please Enter collection name",
            Position=0)]
         [ValidateNotNullOrEmpty()]
            [string[]]$CollectionName,
         
         [Parameter(Mandatory=$false,HelpMessage="Please Enter collection id")]
         [ValidateNotNullOrEmpty()]
            [string[]]$CollectionID
         )
         
    $Class = "SMS_Collection" 
     
     Try {
        $filter = $null
        if($CollectionName) {
            $filter += @("Name LIKE '"+ ($CollectionName -join "' OR Name LIKE '") + "'")
        }
        if ($CollectionID) {
            $filter += @("CollectionID LIKE '"+ ($CollectionID -join "' OR CollectionID LIKE '") + "'")
        }
        
        $filter = $filter -join " AND "
        Get-WmiObject -Namespace "root\SMS\Site_$SiteCode" -Class $Class -Filter $filter -ErrorAction STOP -ComputerName $SiteServer | Select-Object * -ExcludeProperty "__*", "Scope", "Options", "ClassPath", "Properties", "Qualifiers", "Site", "Container"

     }
     Catch {
        Write-Host "Error: $($_.Exception.Message)"
     }
}


Function Get-PatchTuesday {
  [CmdletBinding()]
  Param
  (
    [Parameter(position = 0)]
    [ValidateSet("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
    [String]$weekDay = 'Tuesday',
    [ValidateRange(0, 5)]
    [Parameter(position = 1)]
    [int]$findNthDay = 2
  )
  # Get the date and find the first day of the month
  # Find the first instance of the given weekday
  [datetime]$today = [datetime]::NOW
  $todayM = $today.Month.ToString()
  $todayY = $today.Year.ToString()
  [datetime]$strtMonth = $todayM + '/1/' + $todayY
  while ($strtMonth.DayofWeek -ine $weekDay ) { $strtMonth = $StrtMonth.AddDays(1) }
  $firstWeekDay = $strtMonth

  # Identify and calculate the day offset
  if ($findNthDay -eq 1) {
    $dayOffset = 0
  }
  else {
    $dayOffset = ($findNthDay - 1) * 7
  }
  
  # Return date of the day/instance specified
  $patchTues = $firstWeekDay.AddDays($dayOffset) 
  return $patchTues
}


function Get-ADRInfo
{
    [CMdletbinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$ADRName,
        [Parameter(Mandatory = $true)]
		[string]$SiteCode
    )
    try 
    {
        $Namespace = "root/sms/site_" + $siteCode
        [wmi]$ADR = (Get-WmiObject -Class SMS_AutoDeployment -Namespace $Namespace | Where-Object -FilterScript {$_.Name -eq $ADRName}).__PATH
        return $ADR 
    }
    catch 
    {
        throw 'Failed to Get ADRInfo'
    }
} 

function Set-ADRDeploymentPackage
{
    [CMdletbinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [wmi]$ADRObject,
        [Parameter(Mandatory = $true)]
        [string]$PackageID
        
    )
    try {
        [xml]$ContentTemplateXML = $ADRObject.ContentTemplate
        $ContentTemplateXML.ContentActionXML.PackageID = $PackageID
        $ADRObject.ContentTemplate = $ContentTemplateXML.OuterXml
        $ADRObject.Put() | Out-Null
        Write-Verbose "Succesfully commited updated PackageID"
    }
    catch {
        throw "Something went wrong setting the value"
    }

}



# ----------------------------------------------------------------------------- 
#  MAIN
# ----------------------------------------------------------------------------- 
#Gets the date for patch Tuesday (Midnight)
$patchTuesday = Get-PatchTuesday 

# ----------------------------------------------------------------------------- 
#  Enumerating the times for each collection
# ----------------------------------------------------------------------------- 

If ((Get-Date).IsDaylightSavingTime() -eq $true) {
	$Collection1StartTime = ([DateTime]($PatchTuesday.AddDays($Collection1StartDays).AddHours($Collection1StartHour))).ToString('yyyy/MM/dd HH:mm')
	$Collection1DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection1DeadlineDays).AddHours($Collection1DeadlineHour))).ToString('yyyy/MM/dd HH:mm')

	$Collection2StartTime = ([DateTime]($PatchTuesday.AddDays($Collection2StartDays).AddHours($Collection2StartHour))).ToString('yyyy/MM/dd HH:mm')
	$Collection2DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection2DeadlineDays).AddHours($Collection2DeadlineHour))).ToString('yyyy/MM/dd HH:mm')

	$Collection3StartTime = ([DateTime]($PatchTuesday.AddDays($Collection3StartDays).AddHours($Collection3StartHour))).ToString('yyyy/MM/dd HH:mm')
	$Collection3DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection3DeadlineDays).AddHours($Collection3DeadlineHour))).ToString('yyyy/MM/dd HH:mm')

	$Collection4StartTime = ([DateTime]($PatchTuesday.AddDays($Collection4StartDays).AddHours($Collection4StartHour))).ToString('yyyy/MM/dd HH:mm')
	$Collection4DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection4DeadlineDays).AddHours($Collection4DeadlineHour))).ToString('yyyy/MM/dd HH:mm')

	$Collection5StartTime = ([DateTime]($PatchTuesday.AddDays($Collection5StartDays).AddHours($Collection5StartHour))).ToString('yyyy/MM/dd HH:mm')
	$Collection5DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection5DeadlineDays).AddHours($Collection5DeadlineHour))).ToString('yyyy/MM/dd HH:mm')

	$Collection6StartTime = ([DateTime]($PatchTuesday.AddDays($Collection6StartDays).AddHours($Collection6StartHour))).ToString('yyyy/MM/dd HH:mm')
	$Collection6DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection6DeadlineDays).AddHours($Collection6DeadlineHour))).ToString('yyyy/MM/dd HH:mm')

	$Collection7StartTime = ([DateTime]($PatchTuesday.AddDays($Collection7StartDays).AddHours($Collection7StartHour))).ToString('yyyy/MM/dd HH:mm')
	$Collection7DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection7DeadlineDays).AddHours($Collection7DeadlineHour))).ToString('yyyy/MM/dd HH:mm')

	$Collection8StartTime = ([DateTime]($PatchTuesday.AddDays($Collection8StartDays).AddHours($Collection8StartHour))).ToString('yyyy/MM/dd HH:mm')
	$Collection8DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection8DeadlineDays).AddHours($Collection8DeadlineHour))).ToString('yyyy/MM/dd HH:mm')
	
}


# Adjusts for daylight savings time
If ((Get-Date).IsDaylightSavingTime() -eq $false) {
    $Collection1StartTime = ([DateTime]($PatchTuesday.AddDays($Collection1StartDays).AddHours($Collection1StartHour+1))).ToString('yyyy/MM/dd HH:mm')
    $Collection1DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection1DeadlineDays).AddHours($Collection1DeadlineHour+1))).ToString('yyyy/MM/dd HH:mm')
	
    $Collection2StartTime = ([DateTime]($PatchTuesday.AddDays($Collection2StartDays).AddHours($Collection2StartHour+1))).ToString('yyyy/MM/dd HH:mm')
    $Collection2DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection2DeadlineDays).AddHours($Collection2DeadlineHour+1))).ToString('yyyy/MM/dd HH:mm')

    $Collection3StartTime = ([DateTime]($PatchTuesday.AddDays($Collection3StartDays).AddHours($Collection3StartHour+1))).ToString('yyyy/MM/dd HH:mm')
    $Collection3DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection3DeadlineDays).AddHours($Collection3DeadlineHour+1))).ToString('yyyy/MM/dd HH:mm')

    $Collection4StartTime = ([DateTime]($PatchTuesday.AddDays($Collection4StartDays).AddHours($Collection4StartHour+1))).ToString('yyyy/MM/dd HH:mm')
    $Collection4DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection4DeadlineDays).AddHours($Collection4DeadlineHour+1))).ToString('yyyy/MM/dd HH:mm')

    $Collection5StartTime = ([DateTime]($PatchTuesday.AddDays($Collection5StartDays).AddHours($Collection5StartHour+1))).ToString('yyyy/MM/dd HH:mm')
    $Collection5DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection5DeadlineDays).AddHours($Collection5DeadlineHour+1))).ToString('yyyy/MM/dd HH:mm')

    $Collection6StartTime = ([DateTime]($PatchTuesday.AddDays($Collection6StartDays).AddHours($Collection6StartHour+1))).ToString('yyyy/MM/dd HH:mm')
    $Collection6DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection6DeadlineDays).AddHours($Collection6DeadlineHour+1))).ToString('yyyy/MM/dd HH:mm')

    $Collection7StartTime = ([DateTime]($PatchTuesday.AddDays($Collection7StartDays).AddHours($Collection7StartHour+1))).ToString('yyyy/MM/dd HH:mm')
    $Collection7DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection7DeadlineDays).AddHours($Collection7DeadlineHour+1))).ToString('yyyy/MM/dd HH:mm')

    $Collection8StartTime = ([DateTime]($PatchTuesday.AddDays($Collection8StartDays).AddHours($Collection8StartHour+1))).ToString('yyyy/MM/dd HH:mm')
    $Collection8DeadlineTime = ([DateTime]($PatchTuesday.AddDays($Collection8DeadlineDays).AddHours($Collection8DeadlineHour+1))).ToString('yyyy/MM/dd HH:mm')
}

# ----------------------------------------------------------------------------- 
#  Create new month and/or year folder for Pkg Content
# -----------------------------------------------------------------------------

$YearTest = "E:\Updates\{0}" -f $CurrentYear
$MonthTest = "E:\Updates\{0}\{1}" -f $CurrentYear, $CurrentMonth

If (!(test-path -Path $YearTest))
	{
		New-Item -Itemtype Directory -Force -Path $YearTest
		If (!(test-path -Path $MonthTest))
		{
			New-Item -Itemtype Directory -Force -Path $MonthTest
		}
}
	
Elseif (!(test-path -Path $MonthTest))
	{
		New-Item -Itemtype Directory -Force -Path $MonthTest
}

# ----------------------------------------------------------------------------- 
#  Build New Software Update Package and modify Monthly ADR
# -----------------------------------------------------------------------------
$PackageNameNew = "{0} {1} Security Update Deployment Package" -f $CurrentYear, $CurrentMonth
$PackagePathNew = $UpdateShare + "{0}\{1}" -f $CurrentYear, $CurrentMonth
$PackageDescriptionNew = "Patch Tuesday for {0} of {1}" -f $CurrentMonth, $CurrentYear

$OldADRInfo = Get-ADRInfo -SiteCode $SiteCode -ADRName $ADRName
$NewMonthADR = New-CMSoftwareUpdateDeploymentPackage -Name $PackageNameNew -Path $PackagePathNew -Description $PackageDescriptionNew

Set-ADRDeploymentPackage -ADRObject $OldADRInfo -PackageID $NewMonthADR.PackageID
Start-Sleep -Seconds 60 # Wait 1 minute for ADR info to sync

# ----------------------------------------------------------------------------- 
#  Run ADR
# -----------------------------------------------------------------------------
# Creates Update Group named Monthly Security Updates (Get-Date)
# Creates a deployment with name: ADR for Monthly Security Updates (Get-Date)

Write-Host "Running ADR..."
Get-AutomaticDeploymentRule -Name $ADRName | Start-AutomaticDeploymentRule -Confirm:$false
Write-Host "ADR has been run..."
Start-Sleep -Seconds 5 # Wait 5 second hold for message
Write-Host "Start 5 Minute hold..."
Start-Sleep -Seconds 300 # Wait 5 minutes


# ----------------------------------------------------------------------------- 
#  Update Group
# ----------------------------------------------------------------------------- 

For ($i=0; $i -lt 50; $i++) {
    $UpdateGroup = Get-SoftwareUpdateGroup -Name "$ADRName%"
    If ($UpdateGroup) {
        Write-Host "Update group found!"
        break
    }
    Write-Host "Update group not found after $i attempt(s)..."
    Start-Sleep -Seconds 30 #Wait 30 seconds to show message
}

If (-not $UpdateGroup) {
    throw "Update group was not found after 50 attempts."
}

Write-Host "Renaming update group to $UpdateGroupName"
$UpdateGroup | Rename-SoftwareUpdateGroup -NewName $UpdateGroupName

# ----------------------------------------------------------------------------- 
#  Deployment Creation Phase
# -----------------------------------------------------------------------------

#  Desktop Engineering Pilot
If (![String]::IsNullOrEmpty($Collection1)) {
Write-Host "Creating $Collection1 Deployment..."
$Collection1DisplayName = "1 {0} - {1}" -f $updateGroupName,$Collection1
New-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection1DisplayName -CollectionName $Collection1 -RestartWorkstation $false -RestartServer $false -AvailableDateTime $Collection2StartTime -DeadlineDateTime $Collection2DeadlineTime -SoftwareInstallation $true -AllowRestart $true -UserNotification DisplayAll
Set-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection1DisplayName -Enable $false
}

#  PC Pilot Deployment  
If (![String]::IsNullOrEmpty($Collection2)) {
Write-Host "Creating $Collection2 Deployment..."
$Collection2DisplayName = "2 {0} - {1}" -f $updateGroupName,$Collection2
New-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection2DisplayName -CollectionName $Collection2 -RestartWorkstation $false -RestartServer $false -AvailableDateTime $Collection3StartTime -DeadlineDateTime $Collection3DeadlineTime -SoftwareInstallation $true -AllowRestart $true -UserNotification DisplayAll
Set-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection2DisplayName -Enable $false
}

#  Production PCs Deployment  
If (![String]::IsNullOrEmpty($Collection3)) {
Write-Host "Creating $Collection3 Deployment..."
$Collection3DisplayName = "3 {0} - {1}" -f $updateGroupName,$Collection3
New-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection3DisplayName -CollectionName $Collection3 -RestartWorkstation $false -RestartServer $false -AvailableDateTime $Collection3StartTime -DeadlineDateTime $Collection3DeadlineTime -SoftwareInstallation $true -AllowRestart $true -UserNotification DisplayAll
Set-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection3DisplayName -Enable $false
}

#  MECM Server Deployment  
If (![String]::IsNullOrEmpty($Collection4)) {
Write-Host "Creating $Collection4 Deployment..."
$Collection4DisplayName = "4 {0} - {1}" -f $updateGroupName,$Collection4
New-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection4DisplayName -CollectionName $Collection4 -RestartWorkstation $false -RestartServer $false -AvailableDateTime $Collection4StartTime -DeadlineDateTime $Collection4DeadlineTime -SoftwareInstallation $true -AllowRestart $true -UserNotification DisplayAll
Set-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection4DisplayName -Enable $false
}

#  Server Test/Dev 1 Deployment 
If (![String]::IsNullOrEmpty($Collection5)) {
Write-Host "Creating $Collection5 Deployment..."
$Collection5DisplayName = "5 {0} - {1}" -f $updateGroupName,$Collection5
New-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection5DisplayName -CollectionName $Collection5 -RestartWorkstation $false -RestartServer $false -AvailableDateTime $Collection5StartTime -DeadlineDateTime $Collection5DeadlineTime -SoftwareInstallation $true -AllowRestart $true -UserNotification DisplayAll
Set-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection5DisplayName -Enable $false
}

#  Server Test/Dev 2 Deployment 
If (![String]::IsNullOrEmpty($Collection6)) {
Write-Host "Creating $Collection6 Deployment..."
$Collection6DisplayName = "6 {0} - {1}" -f $updateGroupName,$Collection6
New-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection6DisplayName -CollectionName $Collection6 -RestartWorkstation $false -RestartServer $false -AvailableDateTime $Collection6StartTime -DeadlineDateTime $Collection6DeadlineTime -SoftwareInstallation $true -AllowRestart $true -UserNotification DisplayAll
Set-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection6DisplayName -Enable $false
}

#  Server Test Dev Manual Deployment 
If (![String]::IsNullOrEmpty($Collection7)) {
Write-Host "Creating $Collection7 Deployment..."
$Collection7DisplayName = "7 {0} - {1}" -f $updateGroupName,$Collection7
New-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection7DisplayName -CollectionName $Collection7 -RestartWorkstation $false -RestartServer $false -AvailableDateTime $Collection7StartTime -DeadlineDateTime $Collection7DeadlineTime -SoftwareInstallation $true -AllowRestart $true -UserNotification DisplayAll
Set-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection7DisplayName -Enable $false
}

#  Server Test Dev Manual Deployment 
If (![String]::IsNullOrEmpty($Collection8)) {
Write-Host "Creating $Collection8 Deployment..."
$Collection8DisplayName = "8 {0} - {1}" -f $updateGroupName,$Collection8
New-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection8DisplayName -CollectionName $Collection8 -RestartWorkstation $false -RestartServer $false -AvailableDateTime $Collection8StartTime -DeadlineDateTime $Collection8DeadlineTime -SoftwareInstallation $true -AllowRestart $true -UserNotification DisplayAll
Set-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $UpdateGroupName -DeploymentName $Collection8DisplayName -Enable $false
}

#Distribute content if desired
If ($DistributeConentAuto -eq $true) {
	Start-CMContentDistribution - DeploymentPackageName $PackageNameNew -DistributionPointGroupName $DPGroup
 	Write-Host "Distributing Content for $PackageNameNew!"
  	Start-Sleep -Seconds 30 #Wait 30 seconds to show message
   }
If ($DistributeConentAuto -eq $false) {
 	Write-Host "Content will not be distributed! Please remember to distribute the content for $PackageNameNew!"
  	Start-Sleep -Seconds 30 #Wait 30 seconds to show message
   }

# ----------------------------------------------------------------------------- 
#  Build and Send Email
# ----------------------------------------------------------------------------- 
Write-Host "Sending email..."
$Subject = "Software Updates - {0} {1}" -f $CurrentMonth, $CurrentYear
$Title = "Software Updates for {0} {1}" -f $CurrentMonth, $CurrentYear
$style = "
	body {
        font-family: Calibri
    }
    table {
		width:99%;
		border-top:1px solid #e6e7e9;
		border-right:1px solid #e6e7e9;
		border-collapse:collapse;
	} 
	td {
		color:black;
		border-bottom:1px solid #e6e7e9;
		padding-left:.3em;
		border-left:1px solid #e6e7e9;
		text-align:left;
	}
	th {
		font-weight:normal;
		color: black;
		text-align:left;
		border-bottom: 2px solid #e6e7e9;
		border-left:1px solid #e6e7e9;
		background-color: #A6BCFF;
	}
	th {
		text-align:center;
		font-weight: bold;
		font-size: 1em;
		color:black
	}	
	tfoot th {
		text-align:center;
	}
    h1 {
		font-size: 2em; 
		font-family: `"Segoe UI Light`"; 
		color: white; 
		font-weight: bold; 
		background-color: #3655B3;
		}
	h2 {
		border-bottom: 2px solid #e6e7e9; 
		display:block; 
		padding-left: .5em; 
		padding-top: .3em; 
		padding-bottom:.3em; 
		font-family: `"Segoe UI Light`"; 
		font-weight: bold; font-size: 1.3em
		background-color: #4D4DFF;
		}

    "
$html = "<!DOCTYPE html>`n<html>`n<head>`n`t<title>$title</title>`n<style>$style</style>`n</head>`n<body>`n"
$html += "<h1>$title</h1>`n"
$html += "<h2>Collection Deployments and Deadlines</h2>`n"
$html += Get-CMSoftwareUpdateDeployment -Name $UpdateGroupName | Sort-Object StartTime | Select-Object @{n='Device Collection';e='AssignmentName'},@{n='Available Time';e='StartTime'},@{n='Deadline Time';e='EnforcementDeadline'} | ConvertTo-Html -Fragment
$html += "<h2>Software Updates Deployed to Devices</h2>`n"
$html += Get-CMSoftwareUpdateGroup -Name $UpdateGroupName | Get-CMSoftwareUpdate | Sort-Object LocalizedDisplayName | Select @{n='Article ID';e='ArticleID'},@{n='Patch Name';e='LocalizedDIsplayName'},@{n='Severity';e='SeverityName'} | ConvertTo-Html -Fragment
$html += "`n</body>`n</html>"

Send-MailMessage -SMTPServer $SMTPServerName -To $SCCMTeam -From $SMTPEmailAddress -Subject $Subject -Body $html -BodyAsHTML 
Send-MailMessage -SMTPServer $SMTPServerName -To $ServerTeam -From $SMTPEmailAddress -Subject $Subject -Body $html -BodyAsHTML
Send-MailMessage -SMTPServer $SMTPServerName -To $Management -From $SMTPEmailAddress -Subject $Subject -Body $html -BodyAsHTML
