﻿<#
    .Synopsis
    This script creates functions for general use by other scripts and users
#>

function SetGeneralRoot {
    <#
        .synopsis
        this function sets a script variable for the root folder of General functions.
    #>
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    if (Test-Path "$script:GeneralRoot\GeneralFunctions.ps1") {
        return
    } elseif (Test-Path "$($MyInvocation.PSScriptRoot)\GeneralFunctions.ps1") {
        $script:GeneralRoot = $MyInvocation.PSScriptRoot
    } elseif (Test-Path "$($MyInvocation.InvocationName)\GeneralFunctions.ps1") {
        $script:GeneralRoot = $MyInvocation.InvocationName
    } elseif (Test-Path "$($MyInvocation.PSScriptRoot)\General\GeneralFunctions.ps1") {
        $script:GeneralRoot = "$($MyInvocation.PSScriptRoot)\General"
    } elseif (Test-Path "$($MyInvocation.InvocationName)\General\GeneralFunctions.ps1") {
        $script:GeneralRoot = "$($MyInvocation.InvocationName)\General"
    } elseif (Test-Path "$(Get-Location | select -ExpandProperty Path)\GeneralFunctions.ps1") {
        $script:GeneralRoot = Get-Location | select -ExpandProperty Path
    } elseif (Test-Path "$(Get-Location | select -ExpandProperty Path)\GeneralFunctions.ps1") {
        $script:GeneralRoot = Split-Path (Get-Location | select -ExpandProperty Path)
    } else {
        $f = New-Object System.Windows.Forms.FolderBrowserDialog;
        (New-Object -ComObject Wscript.Shell).Popup("The script needs to determine the root directory in order to proceed. Please choose the directory which contains the GeneralFunctions.ps1 script in the next window.") | Out-Null
        $f.ShowDialog() | Out-Null
        $script:GeneralRoot = $f.SelectedPath
    }
}

function GetFile {
    <#  Opens a dialog box to allow a user to select a file;
        
        Returns: a System.Windows.Forms.OpenFileDialog object
            with a reference to the file;
    #>
    [CmdletBinding()]
    Param(
        [Parameter(HelpMessage='Any warnings you might want to issue to the user. This is optional.')]
        [String]$Prompt
    )

    if (-not ($Prompt -eq $null -or ($Prompt -eq ''))) {(New-Object -ComObject Wscript.Shell).Popup($Prompt) | Out-Null}
    $f = New-Object System.Windows.Forms.OpenFileDialog;
    $f.ShowDialog() | Out-Null
    $f
}

function NewADUserFromExisting {
    Param(
        [Parameter(Mandatory=$True,HelpMessage="The SAM Account Name for this user")]
        [ValidateNotNullOrEmpty()]
        [String]$PreWin2kLogon,

        [Parameter(Mandatory=$True,HelpMessage="The user's first name")]
        [String]$FirstName,

        [Parameter(Mandatory=$True,HelpMessage="The user's last name")]
        [String]$LastName,

        [Parameter(HelpMessage="The user's email address. If not specified, the format will be FirstName.LastName@bahai.org.uk")]
        [String]$Email,

        [Parameter(HelpMessage="The fully qualified domain name for the email address")]
        [String]$Domain,

        [Parameter(Mandatory=$true,HelpMessage="The user's password as a SecureString")]
        [System.Security.SecureString]$Password,

        [Parameter(Mandatory=$true,HelpMessage="The existing user instance on which you will base this user")]
        [ValidateNotNullOrEmpty()]
        [String]$OldUser,

        [Parameter(HelpMessage="The user's role")]
        [String]$RoleTitle,

        [Parameter(HelpMessage="The organization name")]
        [String]$Organization,

        [Parameter(HelpMessage="The user's department")]
        [String]$Department
    )
    
    <# 
        .Synopsis
        This function creates a new AD User by copying the settings from an existing user's profile

    #>

    if ($Email -eq $null -or ($Email -eq '')) {
        $Email = MakeEmail -FirstName $FirstName -LastName $LastName -Domain $Domain
    } else {
        $Email = ValidateEmail -Email $Email
    }
    
    $OUInstance = Get-ADUser -Identity $OldUser -Properties MemberOf,CannotChangePassword,PasswordNeverExpires

    New-ADUser -SamAccountName $PreWin2kLogon `
        -GivenName $FirstName `
        -Surname $LastName `
        -Name "$FirstName $LastName" `
        -DisplayName "$FirstName $LastName" `
        -Title $RoleTitle `
        -Department $Department `
        -Company $Organization `
        -EmailAddress $Email `
        -AccountPassword $Password `
        -UserPrincipalName $Email `
        -Instance $OUInstance `
        -Path ((GetParentOrganizationalUnit -ExistingUser $OUInstance).distinguishedName.ToString())

    ## set the proxyAddress parameter
    Set-ADUser -Identity $PreWin2kLogon -Add @{proxyAddresses="SMTP:$Email"} `
        -CannotChangePassword $OUInstance.CannotChangePassword `
        -PasswordNeverExpires $OUInstance.PasswordNeverExpires

    ## Add user to old user's membership groups
    Select-Object -InputObject $OUInstance -ExpandProperty memberof | ForEach-Object {
        Add-ADGroupMember -Identity ($_.split(',').substring(3)[0]) -Members $PreWin2kLogon
    }
}

function GetParentOrganizationalUnit {
    Param (
        [Parameter(Mandatory=$true,HelpMessage="The existing user account from which you need to draw the Parent OU")]
        [Microsoft.ActiveDirectory.Management.ADUser]$ExistingUser
    )
    
    <#
        .Synopsis
        Use this function to get the DirectoryEntry type of the parent folder for an AD User
    #>

    [ADSI](([ADSI]"LDAP://$($ExistingUser.DistinguishedName)").Parent)
}

function ValidateEmail {
    Param(
        [Parameter(Mandatory=$true,HelpMessage="The user's email address in the format <mailboxName>@<domain>")]
        [ValidateNotNullOrEmpty()]
        [String]$Email
    )
    <#
        .Synopsis
        Validate input in the format <mailboxName>@<domain>.

        .Description
        If only an email is provided, the system will check if it is in the format <mailboxName>@<domain>, where domain will be a string containing at least one dot.
        If an email and Domain are provided, 
    #>
    
    $Email = $Email.ToLower() -replace '\s', ''
    if ($Email -match "^[a-z0-9\-_.]+@[a-z0-9]+\.[a-z0-9]+") {
        $Email
    } else {
        throw "Email parameter must be in the format <mailboxName>@<string1.string2>[.stringN*] and contain only letters, numbers, '@' and/or the '-','_' and '.' punctuation characters"
    }
}

function MakeEmail {
    Param(
        [Parameter(Mandatory=$true,HelpMessage="The user's first name")]
        [ValidateNotNullOrEmpty()]
        [String]$FirstName,
        
        [Parameter(HelpMessage="The user's last name")]
        [String]$LastName, 
        
        [Parameter(Mandatory=$true,HelpMessage="The email domain, in the format <string1.string2[.stringN]*>")]
        [ValidateNotNullOrEmpty()]
        [String]$Domain
    )
    <#
        .Synopsis
        Make an email in the format <firstname>[.lastName]@<domain>
    #>

    $Domain = $Domain.ToLower() -replace '\s',''
    if ($Domain -match "^@?[a-z0-9]+\.[.a-z0-9]+[a-z0-9]$") { 
        
        if (-not ($FirstName -match "^[a-z0-9\-_.]+$")) { 
            throw "FirstName must contain only letters, numbers and/or the '-','_' and '.' punctuation characters"
        }

        $Email = $FirstName.ToLower() -replace '\s',''
            
        if (-not ($LastName -eq $null -or ($LastName -eq ''))) {
            if ($LastName -match "^[a-z0-9\-_.]+$") {
                $Email += '.'+(($LastName.Tolower()) -Replace '\s','')
            } else {
                throw "LastName must contain only letters, numbers and/or the '-','_' and '.' punctuation characters"
            }
        }
            
        if (-not ($Domain[0] -eq '@')) {$Domain = '@'+$Domain}
            
        $Email += $Domain
        $Email

    } else {
        throw 'Parameter Domain must be in the format [@]<string1>.<string2>[.stringN]*'
    }

}


function AssignLicences {
    Param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Email,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Basic","Statistics")]
        [String]$LicenceType
    )
    <#
        .Synopsis
        This function tries to decouple licence assignment from user creation, using a pattern
        similar to the Factory Method to do so.

        .Description
        This function takes as parameter an office 365 user and a string with the licence
        type required, creates the necessary licence options, then assigns them to the user
    #>

    # The Prepare*Licences functions below should be replaced by the relevant 
    # organisation-specific functions to return the Office 365 licences for each,
    # and were not included here because their implementations were too specific
    if ((Get-MsolUser -UserPrincipalName $Email | select -ExpandProperty isLicensed) -eq $false) { 
        switch($LicenceType) {
            "Basic" {$Licences = PrepareBasicLicences; continue}
            "Statistics" {$Licences = PrepareStatisticsLicences; continue}
        }

        $Licences | ForEach-Object { Set-MsolUserLicense -UserPrincipalName $Email `
            -AddLicenses "$($_.AccountSkuId.AccountName):$($_.AccountSkuId.SkuPartNumber)" `
            -LicenseOptions $_ -ErrorAction Continue
        }
    } else {
        Write-Host "User $Email is already licensed. No new licences assigned." -BackgroundColor Black -ForegroundColor Yellow
    }
}

function TestProvision {
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ScriptBlock]$Cmdlet,

        [int]$NumberOfAttempts=30,

        [int]$SecondsBetweenAttempts=60
    )
    <#
        .Synopsis
        This function will test whether a script block's return type is not null 
        once every minute for 30 times, or as many as and for as long as specified 
        on the NumberOfAttempts and SecndsBetweenAttempts parameter. If the script 
        returns something other than $null before the end of the count, the function 
        returns true. Otherwise, it will return false.
    #>
    for ($Count = 1; $Count -le $NumberOfAttempts; $Count++) {
        if ((Invoke-Command -ScriptBlock $Cmdlet) -eq $null) {
            Write-Host "Attempt number $Count " -NoNewline
            Write-Host 'Failed. Will try again in 1 minute.'
            Start-Sleep -Seconds $SecondsBetweenAttempts
        } else {
            return $true
        }
    }
    return $false
}

function SendMessage {
    <# 
     .SYNOPSIS
     This script takes an html message and sends it to the 'To' address in the field
     You need the proper permissions on the mailbox from which you are sending the 
     messages
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,HelpMessage='Your Admin Office 365 credentials')]
        [Alias('Credentials')]
        [System.Management.Automation.PSCredential]$Cred,

        [Parameter(Mandatory=$true,HelpMessage='the path and filename to the message you would like to send')]
        [String]$Path=(GetFile -Prompt "Please choose the html file with the message you would like to send").Filename,

        [Parameter(Mandatory=$true,HelpMessage='The address to which you want to send the message')]
        [String]$To,

        [Parameter(Mandatory=$true,HelpMessage='The email from which you want to send the message')]
        [String]$From,

        [Parameter(Mandatory=$true,HelpMessage='The subject of the email to send out')]
        [String]$Subject,

        [Parameter(HelpMessage='[Optional] If you want it cc''d to anyone')]
        [String[]]$CC
    )


    $Body = ""
    Get-Content -Path $Path | ForEach-Object {$Body += $_}
    Send-MailMessage -To $To -Subject $Subject `
        -From "$From" -Cc $CC -BodyAsHtml:$true `
        -SmtpServer 'smtp.office365.com' -Port 587 -UseSsl:$true `
        -Credential $Cred -Body $Body
    Write-Host "Message sent to $To." -BackgroundColor Black -ForegroundColor Green
}
