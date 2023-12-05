<#
.SYNOPSIS
    This script will help you to perform various Group Policy related tasks
.DESCRIPTION
    This script will help you to perform various Group Policy related tasks
.PARAMETER ReadOnlyMode
    The script will not make any changes in the environment unless this parameter is set to 'False'
.EXAMPLE
    .\GroupPolicyScript.ps1 -ReadOnlyMode $False
.NOTES
    Created by Omer Eldan 
#>

Function Print-GPMainMenu {
    ""
    Write-Host "Welcome to Group Policy Utility (21.4.11)" -ForegroundColor Green 
    "Please select one of the following options:"
    ""
    "(1) - Check total number of GPOs"
    "(2) - Check for disabled GPOs"
    "(3) - Check for unlinked GPOs"
    "(4) - Check for empty GPOs"
    "(5) - Check for GPOs with missing permissions"
    "(6) - Check for GPOs with all links disabled"
    "(7) - Check for disabled GP links"
    "(8) - Get a list of GPOs that are linked only to an empty OU"
    "(9) - Perform a one-time backup of all GPOs"
    "(10) - Create a Schedule Task for Group Policy routine backup"
    "(11) - Get GPO links description"
    "(12) - Create a report with number of affected objects for each GPO"
    "(13) - Create a GPO report"
    "(0) - Exit"
    ""
}

Function Create-OrganizationalUnitsTable {
    $OUs = Get-ADOrganizationalUnit -Filter * -Properties * | select DistinguishedName, CanonicalName
    $OUsArray = @()
    foreach ($OU in $OUs) {
        $OUObject = New-Object -TypeName PSObject
        Add-Member -InputObject $OUObject -MemberType 'NoteProperty' -Name 'DistinguishedName' -Value $OU.DistinguishedName
        Add-Member -InputObject $OUObject -MemberType 'NoteProperty' -Name 'CanonicalName' -Value $OU.CanonicalName
        $OUsArray += $OUObject
    }
    $OUObject = New-Object -TypeName PSObject
    Add-Member -InputObject $OUObject -MemberType 'NoteProperty' -Name 'DistinguishedName' -Value (Get-ADDomain).DistinguishedName
    Add-Member -InputObject $OUObject -MemberType 'NoteProperty' -Name 'CanonicalName' -Value (Get-ADDomain).DNSRoot
    $OUsArray += $OUObject
    Return $OUsArray   
}

Function Get-OrganizationalUnitsDistinguishNameByCanonicalName($CanonicalName) {
    $OUsArray = Create-OrganizationalUnitsTable
    Foreach ($OUObject in $OUsArray) {
        If ($OUObject.CanonicalName -eq $CanonicalName) {
            $DistinguishName = $OUObject.DistinguishedName
            Return $DistinguishName
        }
    }
    Return "NotFound"
}

Function Get-GPTotalGPOs {
    ""
    "Calculating number of GPOs..."
    $TotalGPOs = (Get-GPO -All).Count
    Write-Host "Total GPOs in the environment: $TotalGPOs" -f Yellow
}

Function Get-GPDisabledGPOs ($ReadOnlyMode = $True) {
    ""
    "Looking for disabled GPOs..."
    $DisabledGPOs = @()
    Get-GPO -All | ForEach-Object {
        if ($_.GpoStatus -eq "AllSettingsDisabled") {
            Write-Host "Group Policy " -NoNewline; Write-Host $_.DisplayName -f Yellow -NoNewline; Write-Host " is configured with 'All Settings Disabled'"
            $DisabledGPOs += $_
        }
        Else {
            Write-Host "Group Policy " -NoNewline; Write-Host $_.DisplayName -f Green -NoNewline; Write-Host " is enabled"         
        }
    }
    Write-Host "Total GPOs with 'All Settings Disabled': $($DisabledGPOs.Count)" -f Yellow
    $GPOsToRemove = $DisabledGPOs | Select-Object Id, DisplayName, ModificationTime, GpoStatus | Out-GridView -Title "Showing disabled Group Policies. Select GPOs you would like to delete" -OutputMode Multiple
    if ($ReadOnlyMode -eq $False -and $GPOsToRemove) {
        $GPOsToRemove | ForEach-Object { Remove-GPO -Guid $_.Id -Verbose }
    }
    if ($ReadOnlyMode -eq $True -and $GPOsToRemove) {
        Write-Host "Read-Only mode in enabled. Change 'ReadOnlyMode' parameter to 'False' in order to allow the script make changes" -ForegroundColor Red 
    }
}

Function Get-GPUnlinkedGPOs ($ReadOnlyMode = $True) { 
    ""
    "Looking for unlinked GPOs..."
    $UnlinkedGPOs = @()
    Get-GPO -All | ForEach-Object {
        If ($_ | Get-GPOReport -ReportType XML | Select-String -NotMatch "<LinksTo>" ) {
            Write-Host "Group Policy " -NoNewline; Write-Host $_.DisplayName -f Yellow -NoNewline; Write-Host " is not linked to any object (OU/Site/Domain)"
            $UnlinkedGPOs += $_
        }
        Else {
            Write-Host "Group Policy " -NoNewline; Write-Host $_.DisplayName -f Green -NoNewline; Write-Host " is linked"         
        }
    }
    Write-Host "Total of unlinked GPOs: $($UnlinkedGPOs.Count)" -f Yellow
    $GPOsToRemove = $UnlinkedGPOs | Select-Object -Property Id, DisplayName, ModificationTime | Out-GridView -Title "Showing unlinked Group Policies. Select GPOs you would like to delete" -OutputMode Multiple
    if ($ReadOnlyMode -eq $False -and $GPOsToRemove) {
        $GPOsToRemove | ForEach-Object { Remove-GPO -Guid $_.Id -Verbose }
    }
    if ($ReadOnlyMode -eq $True -and $GPOsToRemove) {
        Write-Host "Read-Only mode in enabled. Change 'ReadOnlyMode' parameter to 'False' in order to allow the script make changes" -ForegroundColor Red 
    }
}

Function Get-GPEmptyGPOs ($ReadOnlyMode = $True) {
    ""
    "Looking for empty GPOs..."
    $EmptyGPOs = @()
    Get-GPO -All | ForEach-Object {
        $IsEmpty = $False
        If ($_.User.DSVersion -eq 0 -and $_.Computer.DSVersion -eq 0) {
            Write-Host "The Group Policy " -nonewline; Write-Host $_.DisplayName -f Yellow -NoNewline; Write-Host " is empty (no settings configured - User and Computer versions are both '0')"
            $EmptyGPOs += $_
            $IsEmpty = $True
        }
        Else {
            [xml]$Report = $_ | Get-GPOReport -ReportType Xml
            If ($Report.GPO.Computer.ExtensionData -eq $NULL -and $Report.GPO.User.ExtensionData -eq $NULL) {
                Write-Host "The Group Policy " -nonewline; Write-Host $_.DisplayName -f Yellow -NoNewline; Write-Host " is empty (no settings configured - No data exist)"
                $EmptyGPOs += $_
                $IsEmpty = $True
            }
        }
        If (-Not $IsEmpty) {
            Write-Host "Group Policy " -NoNewline; Write-Host $_.DisplayName -f Green -NoNewline; Write-Host " is not empty (contains data)"        
        }
    }
    Write-Host "Total of empty GPOs: $($EmptyGPOs.Count)" -f Yellow
    $GPOsToRemove = $EmptyGPOs | Select-Object Id, DisplayName, ModificationTime | Out-GridView -Title "Showing empty Group Policies. Select GPOs you would like to delete" -OutputMode Multiple
    if ($ReadOnlyMode -eq $False -and $GPOsToRemove) {
        $GPOsToRemove | ForEach-Object { Remove-GPO -Guid $_.Id -Verbose }
    }
    if ($ReadOnlyMode -eq $True -and $GPOsToRemove) {
        Write-Host "Read-Only mode in enabled. Change 'ReadOnlyMode' parameter to 'False' in order to allow the script make changes" -ForegroundColor Red 
    }
}

Function Get-GPMissingPermissionsGPOs {
    $MissingPermissionsGPOArray = New-Object System.Collections.ArrayList
    $GPOs = Get-GPO -all
    foreach ($GPO in $GPOs) {
        If ($GPO.User.Enabled) {
            $GPOPermissionForAuthUsers = Get-GPPermission -Guid $GPO.Id -All | Select-Object -ExpandProperty Trustee | ? { $_.Name -eq "Authenticated Users" }
            $GPOPermissionForDomainComputers = Get-GPPermission -Guid $GPO.Id -All | Select-Object -ExpandProperty Trustee | ? { $_.Name -eq "Domain Computers" }
            If (!$GPOPermissionForAuthUsers -and !$GPOPermissionForDomainComputers) {
                $MissingPermissionsGPOArray.Add($GPO) | Out-Null
            }
        }
    }
    If ($MissingPermissionsGPOArray.Count -ne 0) {
        Write-Warning  "The following Group Policy Objects do not grant any permissions to the 'Authenticated Users' or 'Domain Computers' groups:"
        foreach ($GPOWithMissingPermissions in $MissingPermissionsGPOArray) {
            Write-Host "'$($GPOWithMissingPermissions.DisplayName)'"
        }
    }
    Else {
        Write-Host "All Group Policy Objects grant required permissions. No issues were found." -ForegroundColor Green
    }
}

Function Get-GPAllLinksDisabledGPOs {
    
    ""
    "Looking for GPOs with all links set to disabled..."
    $GPOs = Get-GPO -all
    $Counter = 0
    foreach ($GPO in $GPOs) {
        [xml]$Report = Get-GPOReport -Name $GPO.DisplayName -ReportType Xml
        $Links = $Report.GPO.LinksTo
        if ($Links.Count -eq 0) {
            #GPO has no links
            break
        }
        $GPOHasLinkEnabled = $false
        foreach ($Link in $Links) {
            if ($Link.Enabled -eq "true") {
                $GPOHasLinkEnabled = $true
                break
            }
        }
        if ($GPOHasLinkEnabled -eq $false) {
            Write-Host "Group Policy " -NoNewline; Write-Host $GPO.DisplayName -f Yellow -NoNewline; Write-Host " has all links set to disabled"
            $Counter++
        }
    }
    Write-Host "Total GPOs with all links set to disabled: $Counter" -f Yellow
} 


Function Get-GPDisabledGPLinks {
    
    ""
    "Looking for disabled GP links..."
    $GPOs = Get-GPO -all
    $Counter = 0
    foreach ($GPO in $GPOs) {
        [xml]$Report = Get-GPOReport -Name $GPO.DisplayName -ReportType Xml
        $Links = $Report.GPO.LinksTo
        foreach ($Link in $Links) {
            if ($Link.Enabled -eq "false") {
                Write-Host "Group Policy " -NoNewline; Write-Host $GPO.DisplayName -f Yellow -NoNewline; Write-Host " has a disabled link on: " -NoNewline; Write-Host $Link.SOMPath -f Yellow
                $Counter++
            }
        }
    }
    Write-Host "Total disabled GP links: $Counter" -f Yellow
}


Function Backup-GPAllGPOs {
    
    "Backup all GPOs..."
    $Date = Get-Date -Format "yyyy-MM-dd_hh-mm"
    $BackupDir = "C:\Backup\GPO\$Date"
    if ( -Not (Test-Path -Path $BackupDir ) ) {
        New-Item -ItemType Directory -Path $BackupDir
    }
    $ErrorActionPreference = "SilentlyContinue" 
    Backup-GPO -All -Path $BackupDir
    $NumOfBackups = (Get-ChildItem -Path $BackupDir).Count
    $TotalGPOs = (Get-GPO -All).Count
    Get-GPO -All | Select-Object Id, GpoStatus, CreationTime, ModificationTime, DisplayName | ft -AutoSize | Out-File $BackupDir\GroupPolicySummary.csv
    Write-Host "$NumOfBackups out of $TotalGPOs GPOs were backup to the path $BackupDir" -f Yellow
}

Function Create-GPScheduleBackup {
    $Message = "Please enter the credentials of the user which will run the schedule task"; 
    $Credential = $Host.UI.PromptForCredential("Please enter username and password", $Message, "$env:userdomain\$env:username", $env:userdomain)
    $SchTaskUsername = $credential.UserName
    $SchTaskPassword = $credential.GetNetworkCredential().Password
    $SchTaskScriptCode = '$Date = Get-Date -Format "yyyy-MM-dd_hh-mm"
    $BackupDir = "C:\Backup\GPO\$Date"
    $BackupRootDir = "C:\Backup\GPO"
    if (-Not (Test-Path -Path $BackupDir)) {
        New-Item -ItemType Directory -Path $BackupDir
    }
    $ErrorActionPreference = "SilentlyContinue" 
    Get-ChildItem $BackupRootDir | Where-Object {$_.CreationTime -le (Get-Date).AddMonths(-3)} | Foreach-Object { Remove-Item $_.FullName -Recurse -Force}
    Backup-GPO -All -Path $BackupDir'
    $SchTaskScriptFolder = "C:\Scripts\GPO"
    $SchTaskScriptPath = "C:\Scripts\GPO\GPOBackup.ps1"
    if (-Not (Test-Path -Path $SchTaskScriptFolder)) {
        New-Item -ItemType Directory -Path $SchTaskScriptFolder
    }
    if (-Not (Test-Path -Path $SchTaskScriptPath)) {
        New-Item -ItemType File -Path $SchTaskScriptPath
    }
    $SchTaskScriptCode | Out-File $SchTaskScriptPath
    $SchTaskAction = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument "-ExecutionPolicy Bypass $SchTaskScriptPath"
    $Frequency = "Daily", "Weekly"
    $SelectedFrequnecy = $Frequency | Out-GridView -OutputMode Single -Title "Please select the required frequency"
    Switch ($SelectedFrequnecy) {
        Daily {
            $SchTaskTrigger = New-ScheduledTaskTrigger -Daily -At 1am
        }
        Weekly {
            $Days = "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
            $SelectedDays = $Days | Out-GridView -OutputMode Multiple -Title "Please select the relevant days in which the schedule task will run"
            $SchTaskTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $SelectedDays -At 1am
        }
    }  
    Try {
        Register-ScheduledTask -Action $SchTaskAction -Trigger $SchTaskTrigger -TaskName "Group Policy Schedule Backup" -Description "Group Policy $SelectedFrequnecy Backup" -User $SchTaskUsername -Password $SchTaskPassword -RunLevel Highest -ErrorAction Stop
    }
    Catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "Schedule Task regisration was failed due to the following error: $ErrorMessage" -f Red
    }
}

Function Get-GPLinks {
    $GPOName = Read-Host -Prompt 'Please enter Group Policy display name: '
    ""
    Try {
        $GPO = Get-GPO -Name $GPOName -ErrorAction Stop
    }
    Catch {
        Write-Host "$GPO Group Policy object does not exist" -f Red
        Break
    }
    Write-Host "Getting GPO Links for '$GPOName'..." -f Yellow
    [xml]$Report = Get-GPOReport -Name $GPOName -ReportType Xml
    $Report.GPO.LinksTo | Format-Table
}

Function Get-GPOSLinkedToEmptyOUS {
    $emptyOUs = $nonEmptyOUs = @()
    $GPOsLinkedtoEmptyOUs = @()

    ForEach ($OU in Get-ADOrganizationalUnit -Filter { LinkedGroupPolicyObjects -like "*" }) {
        $objects = $null
        $Objects = Get-ADObject -Filter { ObjectClass -ne 'OrganizationalUnit' } -SearchBase $OU
        if ($objects) {
            #Write-Host "OU: '$($OU.Name)' is not empty"
            $nonEmptyOUs += $OU
        }
        Else {
            #Write-Host "OU: '$($OU.Name)' is empty"
            $emptyOUs += $OU
        }
    }

    ForEach ($OU in $emptyOUs) {
        ForEach ($GPOGuid in $OU.LinkedGroupPolicyObjects) {
            $GPO = Get-GPO -Guid $GPOGuid.substring(4, 36)
            Write-Host "GPO: '$($GPO.DisplayName)' is linked only to an empty OU: " -nonewline; Write-Host "$($OU.Name)'" -f yellow
            if ($GPOsLinkedtoEmptyOUs.GPOid -contains $GPO.id) {
                ForEach ($LinkedGPO in ($GPOsLinkedtoEmptyOUs | Where-Object { $_GPOid -eq $GPO.id })) {
                    $LinkedGPO.EmptyOU = [string[]]$LinkedGPO.EmptyOU + "$($OU.DistinguishedName)"
                }
            }
            else {
                $GPOsLinkedtoEmptyOUs += [PScustomObject]@{
                    GPOName    = $GPO.DisplayName
                    GPOId      = $GPO.id
                    EmptyOU    = $OU.DistinguishedName
                    NonEmptyOU = ''
                }
            }
        }
    }
}

Function Create-GPOReportByNumberOfAffectedObjects {
    $GPOs = Get-GPO -all
    $GPOsArray = @()
    foreach ($GPO in $GPOs) {
        $TotalAffectedObjects = 0
        [xml]$Report = Get-GPOReport -Name $($GPO.DisplayName) -ReportType Xml
        $LinkedOUs = $Report.GPO.LinksTo
        foreach ($LinkedOU in $LinkedOUs) {
            $AffectedUsersCount = 0
            $AffectedComputersCount = 0
            $DistinguishName = Get-OrganizationalUnitsDistinguishNameByCanonicalName($LinkedOU.SOMPath)
            $AffectedUsersCount = Get-ADUser -SearchBase $DistinguishName -SearchScope Subtree -Filter * | Measure-Object
            $AffectedComputersCount = Get-ADComputer -SearchBase $DistinguishName -SearchScope Subtree -Filter * | Measure-Object
            $TotalAffectedObjects += $AffectedUsersCount.Count + $AffectedComputersCount.Count
        }
        $GPObject = New-Object -TypeName PSObject
        Add-Member -InputObject $GPObject -MemberType 'NoteProperty' -Name 'Group Policy Name' -Value $($GPO.DisplayName)
        Add-Member -InputObject $GPObject -MemberType 'NoteProperty' -Name 'Number Of Affected Objects' -Value $TotalAffectedObjects
        $GPOsArray += $GPObject
        Write-Host """$($GPO.DisplayName)"" Group Policy is applied by a total of $TotalAffectedObjects objects"
    }
    $GPOsArray | Out-GridView
}

Function Create-GroupPolicyReport {
    "Creating Group Policy Report..."
    $GroupPolicyReportFolder = "C:\Scripts\GPO\Reports"
    If ( -Not (Test-Path -Path $GroupPolicyReportFolder ) ) {
        New-Item -ItemType Directory -Path $GroupPolicyReportFolder
    }
    $GPOs = Get-GPO -All | Select-Object -Property Id, GpoStatus, CreationTime, ModificationTime, DisplayName
    $GPOReport = @()
    ForEach ($GPO in $GPOs) {
        $GPObject = New-Object -TypeName PSObject
        Add-Member -InputObject $GPObject -MemberType 'NoteProperty' -Name 'Group Policy Name' -Value $($GPO.DisplayName)
        Add-Member -InputObject $GPObject -MemberType 'NoteProperty' -Name 'Group Policy ID' -Value $($GPO.Id)
        Add-Member -InputObject $GPObject -MemberType 'NoteProperty' -Name 'Group Policy Status' -Value $($GPO.GpoStatus)
        Add-Member -InputObject $GPObject -MemberType 'NoteProperty' -Name 'Group Policy Creation Time' -Value $($GPO.CreationTime)
        Add-Member -InputObject $GPObject -MemberType 'NoteProperty' -Name 'Group Policy Modification Time' -Value $($GPO.ModificationTime)
        [xml]$Report = Get-GPOReport -Name $GPO.DisplayName -ReportType Xml
        $Links = $Report.GPO.LinksTo
        $EnabledLinks = ""
        foreach ($Link in $Links) {
            If ($($Link.Enabled -eq 'true')) {
                $EnabledLinks += "$($Link.SOMPath),"
            }
        }
        if ($EnabledLinks) {
            $EnabledLinks = $EnabledLinks.Substring(0, $EnabledLinks.Length - 1)
        }
        Else {
            $EnabledLinks = "GPO is not linked to any object"
        }
        Add-Member -InputObject $GPObject -MemberType 'NoteProperty' -Name 'Group Policy Enabled Links' -Value ($EnabledLinks)
        $GPOReport += $GPObject
    }
    $GPOReport | Export-Csv $GroupPolicyReportFolder\GroupPolicyReport.csv
    $GPOReport | Out-GridView
}

Function Create-AGPMReport {
    $AGPMArchiveLocation = "C:\ProgramData\Microsoft\AGPM"
    $AGPMArchive = Get-ChildItem -Directory -Path  $AGPMArchiveLocation
    $GPOArray = @()
    $GPOGUIDArray
    $IgnoreAGPMVersions = $true
    Foreach ($ArchivedGPO in $AGPMArchive) {
        Try {
            [xml]$GPReport = Get-Content -Path "$($ArchivedGPO.FullName)\gpreport.xml" -ErrorAction Stop
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            Write-Error "Failed to get gpreport.xml in $($ArchivedGPO.FullName). The following error has occurred: $ErrorMessage"
            $GPOObject = New-Object -TypeName psobject
            Add-Member -InputObject $GPOObject -MemberType 'NoteProperty' -Name 'Status' -Value "Error - AGPM archive is damaged"
            Add-Member -InputObject $GPOObject -MemberType 'NoteProperty' -Name 'GPO-AGPM-GUID' -Value $ArchivedGPO.Name
            Add-Member -InputObject $GPOObject -MemberType 'NoteProperty' -Name 'GPO-SYSVOL-GUID' -Value "No information"
            $GPOArray += $GPOObject
            Continue
        }
        $GPOGUID = $GPReport.GPO.Identifier.Identifier.'#text'
        if ($IgnoreAGPMVersions -and $GPOGUIDArray.Contains($GPOGUID)) {
            Write-Host "Group Policy archive was already checked. Skip."
            Continue
        }
        $GPOGUIDArray = + $GPOGUID
        $GPO = Get-GPO -Guid $GPOGUID | Select-Object -Property Id, DisplayName, GpoStatus, CreationTime, ModificationTime -ErrorAction Continue
        if (!$GPO) {
            $ErrorMessage = $_.Exception.Message
            Write-Error "Failed to get Group Policy Object with GUID $GPOGUID. The following error has occurred: $ErrorMessage"
            $GPOObject = New-Object -TypeName psobject
            Add-Member -InputObject $GPOObject -MemberType 'NoteProperty' -Name 'Status' -Value "Error - GPO does not exist in production (SYSVOL)"
            Add-Member -InputObject $GPOObject -MemberType 'NoteProperty' -Name 'GPO-AGPM-GUID' -Value $ArchivedGPO.Name
            Add-Member -InputObject $GPOObject -MemberType 'NoteProperty' -Name 'GPO-SYSVOL-GUID' -Value "Not exist"
            $GPOArray += $GPOObject
        }
        else {
            $GPOObject = New-Object -TypeName psobject
            Add-Member -InputObject $GPOObject -MemberType 'NoteProperty' -Name 'Status' -Value "Valid - GPO exist in both archive (AGPM) and production (SYSVOL)"
            Add-Member -InputObject $GPOObject -MemberType 'NoteProperty' -Name 'GPO-AGPM-GUID' -Value $ArchivedGPO.Name
            Add-Member -InputObject $GPOObject -MemberType 'NoteProperty' -Name 'GPO-SYSVOL-GUID' -Value $GPO.Id
            $GPOArray += $GPOObject
        }
    }
    $GPOArray | Out-GridView -Title "Group Policies Summary - SYSVOL and AGPM"
}

Clear-Host
Import-Module GroupPolicy
$Exit = 0
while ($Exit -ne 1) {
    Print-GPMainMenu
    $Selection = Read-Host -Prompt 'Please enter your choice'
    Switch ($Selection) {
        1 { Get-GPTotalGPOs }
        2 { Get-GPDisabledGPOs }
        3 { Get-GPUnlinkedGPOs }
        4 { Get-GPEmptyGPOs }
        5 { Get-GPMissingPermissionsGPOs }
        6 { Get-GPAllLinksDisabledGPOs }
        7 { Get-GPDisabledGPLinks }
        8 { Get-GPOSLinkedToEmptyOUS }
        9 { Backup-GPAllGPOs }
        10 { Create-GPScheduleBackup }
        11 { Get-GPLinks }
        12 { Create-GPOReportByNumberOfAffectedObjects }
        13 { Create-GroupPolicyReport }
        0 { $Exit = 1 }
        default { 'Unknown selection. Please select a number from the menu' }
    }
}
Clear-Host