<# 
    .Synopsis
    This script contains functions for login purposes
#>
[CmdletBinding()]
Param()

function LoginO365 {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,HelpMessage='Your Office 365 Admin credentials')]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCredential] $Cred,

        [Parameter(Mandatory=$true,HelpMessage='Your Sharepoint Online Admin site URL')]
        [ValidateNotNullOrEmpty()]
        [String]$SPAdminSite
    )

    ConnectToExchangeOnline $Cred
    ConnectToMSOnline $Cred
    ConnectToSharepointOnline -Cred $Cred -SPAdminSite $SPAdSite
}


function ModuleChecker {
    # This function checks if a module is installed
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,HelpMessage='The name of the module as you would expect to see on get-module')]
        [Alias('ModuleName')]
        [String]$ModName,

        [Parameter(HelpMessage='If true, most messages will be supressed. Use this if you want less verbose output')]
        [Boolean]$SuppressMessages=$false
    )

    # setting error action
    $ErrorActionPreference = 'Stop'

    if (Get-Module -Name $ModName) {
        if (-not $SuppressMessages) {
            Write-Host "`n$ModName Module loaded." -BackgroundColor Black -ForegroundColor Green
        }
    } elseif (Get-Module -ListAvailable -Name $ModName) {
        if (-not $SuppressMessages) {Write-Host "`nLoading $ModName Module..." -NoNewline -BackgroundColor Black -ForegroundColor Cyan}
        Import-Module -Name $ModName 3> $null | Out-Null
        if (-not $SuppressMessages) {Write-Host "Done!" -BackgroundColor Black -ForegroundColor Green}
    } else {
        $ModName = $ModName+' module not found. Please download and/or install it before proceeding'
        throw $ModName
    }
}

function ConnectToExchangeOnline {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,HelpMessage="Your Office 365 admin credentials")]
        [Alias('Credentials')]
        [System.Management.Automation.PSCredential]$Cred
    )

    # setting error action
    $ErrorActionPreference = 'Stop'

    # Check if Exchange Online Session already exists
    $Session = Get-PSSession | Where-Object {$_.ConfigurationName -eq 'Microsoft.Exchange'}

    try {
        if (-not ($Session.State -eq 'Opened')) { # this will also be true if $Session is $null
            if (-not ($Session -eq $null)) {
                Remove-PSSession $Session
            }
            Write-Host 'Attempting to connect to Microsoct Exchange Online... ' -NoNewline -BackgroundColor Black -ForegroundColor Cyan
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ `
                -Credential $Cred -Authentication Basic –AllowRedirection 3> $null
            Import-PSSession $Session -AllowClobber 3>$null | Out-Null
            Write-Host 'Done!' -BackgroundColor Black -ForegroundColor Green
            Write-Host 'Connected to Microsoft Exchange Online' -BackgroundColor Black -ForegroundColor Green
        } else {
            Write-Host 'Already connected to Microsoft Exchange Online' -BackgroundColor Black -ForegroundColor Green
        }
    } catch {
        Write-Warning 'Could not connect to Exchange Online at this time. Please check your connection setting and/or the instructions on the script.'
        throw $_
    }
}

function ConnectToMSOnline {
    # This script connects you to MSOnline (Azure Active Directory)
    
    [CmdletBinding()]

    Param (
        [Parameter(Mandatory=$true,HelpMessage='Your Admin Office 365 credentials')]
        [Alias('Credentials')]
        [System.Management.Automation.PSCredential]$Cred
    )
    # setting error action
    $ErrorActionPreference = 'Stop'

    # Import MSOnline Module
    try {

        ModuleChecker -ModName MSOnline

    } catch {
        Write-Warning "$($Path): Could not load ModuleChecker.ps1. Please change the path if the file has been moved"
        throw $_
    }

    # Connect to MSOnline
    try {
        Write-Host 'Attempting to connect to AzureAD (MSOnline)... ' -NoNewline -BackgroundColor Black -ForegroundColor Cyan
        Get-MsolDomain -ErrorAction Stop | Out-Null
    } catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException] {
        Connect-MsolService -Credential $Cred
    }
    Write-Host 'Connected!' -BackgroundColor Black -ForegroundColor Green
    Write-Host
}

function ConnectToSharepointOnline {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,HelpMessage='Your Admin Office 365 credentials')]
        [Alias('Credentials')]
        [System.Management.Automation.PSCredential]$Cred,

        [Parameter(HelpMessage='The Sharepoint Admin site address.',Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$SPAdminSite,

        [Parameter(HelpMessage='If true, most messages will be supressed. Use this if you want less verbose output')]
        [Boolean]$SuppressMessages=$false
    )
    # Add SPOnline Module
    try {
        if ($MyInvocation.PSScriptRoot -eq $null) {
            $script:ScriptPath = Split-Path $MyInvocation.InvocationName -Parent
        } else {
            $script:ScriptPath = $MyInvocation.PSScriptRoot
        }

        . $script:ScriptPath\GeneralFunctions.ps1

        ModuleChecker -ModName 'Microsoft.Online.SharePoint.PowerShell' -SuppressMessages $SuppressMessages
    } catch {
        throw $_
    }

    # Connect to Sharepoint Online
    try {
        if(-not $SuppressMessages) {
            Write-Host 'Attempting to connect to Sharepoint Online... ' -NoNewline `
                -BackgroundColor Black -ForegroundColor Cyan
        }
        Get-SPOSite | Out-Null
    } catch [System.InvalidOperationException] {
        Connect-SPOService -Url $SPAdminSite -Credential $Cred
    }
    if(-not $SuppressMessages) {Write-Host 'Connected!' -BackgroundColor Black -ForegroundColor Green}
}