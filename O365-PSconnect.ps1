#Requires -RunAsAdministrator
#Requires -Modules MsOnline, Microsoft.Online.SharePoint.PowerShell, SkypeOnlineConnector

# more details at https://technet.microsoft.com/en-us/library/dn568015.aspx

#region FUNCTIONS

<#
.Synopsis
   Display dynamic menu of O365 services to connect to based on $ObjArr members.  Menu prompt will loop until 'c' 
   is or 'All' is selected at which point the script will attempt to connect to O365 services.  Basic input validation
   will take place for menu selections
#>
function Display-Menu
{
    [CmdletBinding()]
    Param([Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$ObjArr)

    $index = 1  # anchor at 1 otherwise regex comparison may fail
    $lowerbound = $index
    $upperbound = $ObjArr.count + $index
    
    $hash = @{}
    $ObjArr | % {$hash[$index] = $_.description; $index++} # hash used to map numeric menu selections to description
    $hash[$index] = 'All' # utilize last index increment from loop above
        
    $selections = @()
    $userInput = ''
    
    do {
        write-host "`n"
        if($userInput -ne '' -and $userInput -notlike "[$lowerbound-$upperbound]") {Write-Host "`nInvalid Selection: $userInput`n" -f Yellow}
        
        $hash.GetEnumerator() | sort name | % {Write-Host "$($_.name.tostring())) $($_.value)"}
      
        $userInput = Read-Host "Enter O365 service to connect to (c to start connections)"
        if($userInput -like "[$lowerbound-$upperbound]") {$selections += $userInput}

    } while ($userInput -notmatch '^c$' -and $userInput -ne $upperbound)

    $selections | sort | select -Unique
}

<#
.Synopsis
   Attempt to connect to an O365 service by importing modules or PSsessions based on the $O365service object
#>
function Connect-O365service
{
    [CmdletBinding()]
    Param([Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$O365service)
    
    Write-Host "`n"
    if ($O365service.ImportString) { 
        try 
        {
            Write-Host "Importing $($O365service.Description) module" -f Green
            Invoke-Expression $O365service.ImportString -ea Stop
        } 
        catch {Write-Host $_.exception.message -f Red} 
    }

    if ($O365service.ConnectionString) { 
        try 
        {
            Write-Host "Connecting to $($O365service.Description) as $($global:O365cred.UserName)" -f Green
            Invoke-Expression $O365service.ConnectionString -ea Stop | Out-Null
        } 
        catch {Write-Host $_.exception.message -f Red}
    }
}

#endregion FUNCTIONS

#region MAIN

# Basic validation done for O365 admin credential format
$global:O365cred = $null
do {$global:O365cred = Get-Credential -Message "Enter O365 Admin Credentials"; if(!$global:O365cred){exit}} while ($global:O365cred.UserName -notmatch '@[\w]+\.[\w]+' -and $global:O365cred)

# This can be hardcoded or use read-host if admin UPN differs from tenant FQDN
# See SharePoint menu item for usage
$sharepointPrefix = $global:O365cred.UserName -match "(?<=@)[\w]+" | % {$matches[0]} 

# Create custom objects to be used for dynamic menu display
# New menu objects should have Description, ImportString, and ConnectionString attributes
$customMenuArr = @()

# Microsoft Online 
$propertySet = @{Description = 'Microsoft Online'; ImportString = 'Import-Module MsOnline'; ConnectionString = 'Connect-MsolService -Credential $global:O365cred'}
$customMenuArr += New-Object -TypeName PSobject -Property $propertySet

# Exchange
$propertySet = @{Description = 'Exchange Online'; ImportString = ''; ConnectionString = 'Import-PSSession -Session (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $global:O365cred -Authentication "Basic" -AllowRedirection) -DisableNameChecking'}
$customMenuArr += New-Object -TypeName PSobject -Property $propertySet

# SharePoint
$propertySet = @{Description = 'SharePoint Online'; ImportString = 'Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking'; ConnectionString = "Connect-SPOService -Url https://$sharepointPrefix-admin.sharepoint.com -credential `$global:O365cred"}
$customMenuArr += New-Object -TypeName PSobject -Property $propertySet

# Skype for Business
$propertySet = @{Description = 'Skype for Business Online'; ImportString = 'Import-Module SkypeOnlineConnector'; ConnectionString = 'Import-PSSession -Session (New-CsOnlineSession -Credential $global:O365cred)'}
$customMenuArr += New-Object -TypeName PSobject -Property $propertySet

# Security and Compliance Center
$propertySet = @{Description = 'Security and Compliance Center'; ImportString = ''; ConnectionString = 'Import-PSSession -Session (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $global:O365cred -Authentication Basic -AllowRedirection) -Prefix cc'}
$customMenuArr += New-Object -TypeName PSobject -Property $propertySet


$menuSelections = Display-Menu -ObjArr $customMenuArr

write-host "`n"
if($menuSelections -eq $null) {
    Write-Host "No services selected" -f Cyan
} elseif ($menuSelections -contains ($customMenuArr.Count + 1).ToString()) {
    Write-Host "All services selected" -f Cyan
    $customMenuArr | % {Connect-O365service -O365service $_}
} else {
    Write-Host ($menuSelections | % {$customMenuArr[$_-1].description} | Out-String).Replace("`n",', ').trim(', ') "services selected" -f Cyan
    $menuSelections | % {Connect-O365service -O365service $customMenuArr[$_-1]}
}

#endregion MAIN