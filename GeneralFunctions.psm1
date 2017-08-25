<#
    .Synopsis
    This script creates functions for general use by other scripts and users
#>

function SetGeneralRoot {
    <#
        .synopsis
        this function sets a script variable for the root folder of General functions.
    #>
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    if (Test-Path "$global:GeneralRoot\GeneralFunctions.psm1") {
        return
    } elseif (Test-Path "$($MyInvocation.PSScriptRoot)\GeneralFunctions.psm1") {
        $global:GeneralRoot = $MyInvocation.PSScriptRoot
    } elseif (Test-Path "$($MyInvocation.InvocationName)\GeneralFunctions.psm1") {
        $global:GeneralRoot = $MyInvocation.InvocationName
    } elseif (Test-Path "$($MyInvocation.PSScriptRoot)\General\GeneralFunctions.psm1") {
        $global:GeneralRoot = "$($MyInvocation.PSScriptRoot)\General"
    } elseif (Test-Path "$($MyInvocation.InvocationName)\General\GeneralFunctions.psm1") {
        $global:GeneralRoot = "$($MyInvocation.InvocationName)\General"
    } elseif (Test-Path "$(Get-Location | select -ExpandProperty Path)\GeneralFunctions.psm1") {
        $global:GeneralRoot = Get-Location | select -ExpandProperty Path
    } elseif (Test-Path "$(Get-Location | select -ExpandProperty Path)\GeneralFunctions.psm1") {
        $global:GeneralRoot = Split-Path (Get-Location | select -ExpandProperty Path)
    } else {
        $f = New-Object System.Windows.Forms.FolderBrowserDialog;
        (New-Object -ComObject Wscript.Shell).Popup("The script needs to determine the root directory in order to proceed. Please choose the directory which contains the GeneralFunctions.psm1 script in the next window.") | Out-Null
        $f.ShowDialog() | Out-Null
        $global:GeneralRoot = $f.SelectedPath
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

function LoadModuleIntoSession {
    <#
        .synopsis
        This function takes an open session (with the machine which has the ActiveDirectory module) 
        and loads the ActiveDirectory and GeneralFunctions modules into it. It returns the session
        with the modules now loaded.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,HelpMessage="An open session with the server running the Active Directory")]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.Runspaces.PSSession]$ADSession
    )
    Invoke-Command -Session $ADSession -ScriptBlock {Import-Module -Name ActiveDirectory}

    $ScriptContents = Get-Content -Path "$Global:GeneralRoot\GeneralFunctions.psm1"
    $ModuleName = 'GeneralFunctions'
    $FileName = "$ModuleName.psm1"
    
    Invoke-Command -Session $ADSession -ScriptBlock {
        Param($ScriptContents,$FileName)
        Set-Content -Path $FileName -Value $ScriptContents -Force
        Import-Module -Name ".\$FileName"
    } -ArgumentList ($ScriptContents,$FileName)

    $ADSession
}

function NewADUserFromExisting {
    <# 
        .Synopsis
        This function creates a new AD User by copying the settings from an existing user's profile

    #>
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
    

    if ($Email -eq $null -or ($Email -eq '')) {
        $Email = MakeEmailAddress -FirstName $FirstName -LastName $LastName -Domain $Domain
    } else {
        $Email = ValidateEmailAddress -Email $Email
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
        $Group = $_.split(',').substring(3)[0]    
        if (-not (Get-ADGroupMember -Identity $Group | Select-Object -ExpandProperty SamAccountName).Contains($PreWin2kLogon)) {
            Add-ADGroupMember -Identity $Group -Members $PreWin2kLogon -ErrorAction Continue
        }
    } -ErrorAction Continue
}


function NewADUserFromExistingWithHashTable {
    <#
        .Synopsis
        This function takes in a HashTable loaded with new users' details and creates them in the active directory
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,HelpMessage='A HashTable loaded with the details of the new users')]
        [PSCustomObject]$HashTable,

        [Parameter(Mandatory,HelpMessage='The name of your organization')]
        [String]$Organization
    )


    $HashTable | ForEach-Object {
        NewADUserFromExisting -PreWin2kLogon $_.ADLogonID `
            -FirstName $_.FirstName `
            -LastName $_.LastName `
            -Email $_.Email `
            -Password (ConvertTo-SecureString -AsPlainText $_.ADPassword -Force) `
            -OldUser $_.CloneFromUser `
            -RoleTitle $_.Role `
            -Organization $Organization `
            -Department $_.Department `
    }
}

function NewADUserFromExistingWithCsv {
    <#
        .Synopsis
        This function takes the new user parameters from a file and calls the 
        NewADUserFromExisting-WithHashTable function on it
    #>

    [CmdletBinding()]
    Param(
        [Parameter(HelpMessage='The path to the csv file with the user''s details')]
        [String]$File=(GetFile),

        [Parameter(Mandatory,HelpMessage='The name of your organization')]
        [String]$Organization
    )

    NewADUserFromExistingWithHashTable -HashTable (Import-Csv -Path $File) -Organization $Organization
}


function GetParentOrganizationalUnit {
    <#
        .Synopsis
        Use this function to get the DirectoryEntry type of the parent folder for an AD User
    #>
    Param (
        [Parameter(Mandatory=$true,HelpMessage="The existing user account from which you need to draw the Parent OU")]
        [Microsoft.ActiveDirectory.Management.ADUser]$ExistingUser
    )
    

    [ADSI](([ADSI]"LDAP://$($ExistingUser.DistinguishedName)").Parent)
}

function ValidateEmailAddress {
    <#
        .Synopsis
        Validate input in the format <mailboxName>@<domain>.

        .Description
        If only an email is provided, the system will check if it is in the format <mailboxName>@<domain>, where domain will be a string containing at least one dot.
        If an email and Domain are provided, 
    #>
    
    Param(
        [Parameter(Mandatory=$true,HelpMessage="The user's email address in the format <mailboxName>@<domain>")]
        [ValidateNotNullOrEmpty()]
        [String]$Email
    )
    $Email = $Email.ToLower() -replace '\s', ''
    if ($Email -match "^[a-z0-9\-_.]+@[a-z0-9]+\.[a-z0-9]+") {
        $Email
    } else {
        throw "Email parameter must be in the format <mailboxName>@<string1.string2>[.stringN*] and contain only letters, numbers, '@' and/or the '-','_' and '.' punctuation characters"
    }
}

function MakeEmailAddress {
    <#
        .Synopsis
        Make an email address in the format <firstname>[.lastName]@<domain>
    #>

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
    <#
        .Synopsis
        This function tries to decouple licence assignment from user creation, using a pattern
        similar to the Factory Method to do so.

        .Description
        This function checks if the user is already licensed and, if not, takes as parameter 
        an office 365 user and a string with the licence type required, creates the necessary 
        licence options, then assigns them to the user. Checking if user is licensed can be
        bypassed by providing the relevant argument.
    #>
    Param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Email,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Full","SharepointOnly")]
        [String]$LicenceType,

        [Parameter(HelpMessage="Set to true if you want to override any possible licence the user may already have")]
        [Boolean]$Override=$false
    )

    # The Prepare*Licences functions below should be replaced by the relevant 
    # organisation-specific functions to return the Office 365 licences for each,
    # and were not included here because their implementations were too specific
    if ((Get-MsolUser -UserPrincipalName $Email | select -ExpandProperty isLicensed) -eq $false `
        -or $Override) { 
        switch($LicenceType) {
            "Full" {$Licences = PrepareFullLicences; continue}
            "SharepointOnly" {$Licences = PrepareSharepointOnlyLicences; continue}
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
    <#
        .Synopsis
        This function will test whether a script block's return type is not null 
        once every minute for 30 times, or as many as and for as long as specified 
        on the NumberOfAttempts and SecndsBetweenAttempts parameter. If the script 
        returns something other than $null before the end of the count, the function 
        returns true. Otherwise, it will return false.
    #>
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ScriptBlock]$Cmdlet,

        [int]$NumberOfAttempts=30,

        [int]$SecondsBetweenAttempts=60
    )
    for ($Count = 1; $Count -le $NumberOfAttempts; $Count++) {
        if ((Invoke-Command -ScriptBlock $Cmdlet) -eq $null) {
            Write-Host "`nAttempt number $Count " -NoNewline
            Write-Host "Failed. Will try again in $SecondsBetweenAttempts seconds." -NoNewline
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
    [CmdletBinding(DefaultParameterSetName='Path')]
    Param(
        [Parameter(Mandatory, HelpMessage='Your Admin Office 365 credentials',ParameterSetName='Path')]
        [Parameter(Mandatory, ParameterSetName='Msg')]
        [Alias('Credentials')]
        [System.Management.Automation.PSCredential]$Cred,

        [Parameter(Mandatory, ParameterSetName='Path',
            HelpMessage='the path and filename to the message you would like to send')]
        [String]$Path<#=(GetFile -Prompt "Please choose the file with the message you would like to send").Filename#>,

        [Parameter(Mandatory, HelpMessage="A string containing the message which you want to send",
            ParameterSetName='Msg')]
        [String]$Msg,

        [Parameter(Mandatory,
            HelpMessage='The address to which you want to send the message',
            ParameterSetName='Path')]
        [Parameter(Mandatory,
            ParameterSetName='Msg')]
        [String[]]$To,

        [Parameter(Mandatory,
            HelpMessage='The email from which you want to send the message',
            ParameterSetName='Path')]
        [Parameter(Mandatory,
            ParameterSetName='Msg')]
        [String]$From,

        [Parameter(Mandatory, HelpMessage='The subject of the email to send out',
            ParameterSetName='Path')]
        [Parameter(Mandatory,ParameterSetName='Msg')]
        [ValidateNotNullOrEmpty()]
        [String]$Subject,

        [Parameter(HelpMessage='[Optional] If you want it cc''d to anyone',
            ParameterSetName='Path')]
        [Parameter(ParameterSetName='Msg')]
        [String[]]$CC
    )


    $Body = ""
    if ($PSCmdlet.ParameterSetName -eq 'Path') {
        Get-Content -Path $Path | ForEach-Object {$Body += $_}
    } else {
        $Body = $Msg
    }

    if ($CC.Count -eq 0 -or ($CC[0] -eq '')) {
        Send-MailMessage -To $To -Subject $Subject `
            -From "$From" -BodyAsHtml:$true `
            -SmtpServer 'smtp.office365.com' -Port 587 -UseSsl:$true `
            -Credential $Cred -Body $Body -Encoding UTF8
    } else {
        Send-MailMessage -To $To -Subject $Subject `
            -From "$From" -Cc $CC -BodyAsHtml:$true `
            -SmtpServer 'smtp.office365.com' -Port 587 -UseSsl:$true `
            -Credential $Cred -Body $Body -Encoding UTF8
    }

    Write-Host "Message sent to $To." -BackgroundColor Black -ForegroundColor Green
}

function MakeHtmlFromTemplate {
    <#
        .Synopsis
        This function takes in a template with placeholders, replaces the placeholders with 
        arguments provided, then returns html output.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(HelpMessage="The path to the file containing the message template")]
        [String]$PathToTemplate=(GetFile -Prompt 'Please find the file containing the message template. Make sure that all the placeholders are in the format "\$_[a-z]+"').Filename,

        [Parameter(Mandatory,HelpMessage="A hash table containing the placeholder names and the values by which they should be replaced")]
        [System.Collections.Hashtable]$Values
    )

    $Body = ""
    Get-Content -Path $PathToTemplate | ForEach-Object {$Body += $_}
    $Values.GetEnumerator() | ForEach-Object {$Body = $Body.Replace($_.key,$_.value)}
    $Body
}

function GiveMailboxPermissionsWithSMA {
    <#
        .Synopsis
        This function creates an SMA, adds members to it, then gives both 
        Send As and Full Access permission to the selected mailbox by 
        adding the SMA to it's permissions and recipient permissions.
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,HelpMessage='The name of the security group (e.g. SMA <Mailbox>)')]
        [String]$SMAName,
        
        [Parameter(Mandatory,HelpMessage="The email domain, in the format <string1.string2[.stringN]*>")]
        [ValidateNotNullOrEmpty()]
        [String]$Domain,

        [Parameter(Mandatory,HelpMessage="The member or members who will be part of the SMA")]
        [ValidateNotNullOrEmpty()]
        [String[]]$SMAMembers,
        
        [Parameter(Mandatory,HelpMessage="The mailbox to which the users need to gain access")]
        [ValidateNotNullOrEmpty()]
        [String]$Mailbox,

        [Parameter(HelpMessage="If the SMA should have ReadOnly access, set to true. Default false")]
        [ValidateNotNullOrEmpty()]
        [Boolean]$ReadOnly=$false
    )
    if (-not ($SMAName.ToLower().Contains("sma"))) {
        $SMAName = "SMA $SMAName"
    }

    $Alias = $SMAName.Replace(" ","")
    $SMAEmail = MakeEmailAddress -FirstName $Alias -Domain $Domain

    New-DistributionGroup -Name $SMAName -Alias $Alias -PrimarySmtpAddress $SMAEmail -Members $SMAMembers -Type 'Security'
    Set-DistributionGroup -Identity $SMAEmail -HiddenFromAddressListsEnabled $true
    Add-MailboxPermission -Identity $Mailbox -AccessRights FullAccess -User $SMAName
    if (-not $ReadOnly) {
        Add-RecipientPermission -Identity $Mailbox -AccessRights SendAs -Trustee $SMAEmail
    }
}


function TestSPAccess {
    <#
        .Synopsis
        This function takes in a "legitimate" and "test" users' email addresses and adds
        the test user to exactly the same groups as the legitimate one. This allows you to
        test what exactly the legitimate user can see on Sharepoint to make sure that
        they have the righ access.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(HelpMessage='The legitimate user''s O365 email')]
        [String]$User=(Read-Host -Prompt 'Please enter the user''s O365 email address'),

        [Parameter(HelpMessage='The test user''s O365 email')]
        [String]$TestUser=(Read-Host -Prompt 'Please enter the user''s O365 email address'),

        [Parameter(Mandatory=$true,HelpMessage="The URL for your Sharepoint team site")]
        [ValidateNotNullOrEmpty()]
        [String] $SPSite

    )

    # remove test user from all groups it is currently in
    Get-SPOUser -Site $SPSite -LoginName $TestUser | select -ExpandProperty groups | % {Remove-SPOUser -Site $SPSite -LoginName $TestUser -Group $_}

    # add test user to the groups of the chosen user
    Get-SPOUser -Site $SPSite -LoginName $User | select -ExpandProperty groups | % {Add-SPOUser -Site $SPSite -LoginName $TestUser -Group $_}
}


function GetUserDGMembership {
    <#
        .synopsis
        This function returns all the distribution groups of which a user is a member
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,HelpMessage='The user''s Office 365 account')]
        [String]$UserEmail
    )

    Get-DistributionGroup -ResultSize Unlimited | Where-Object {
        Get-DistributionGroupMember -Identity $_.PrimarySmtpAddress -ResultSize Unlimited | Where-Object {
            $_.PrimarySmtpAddress -eq $UserEmail
        }
    }
}

function BlockOffice365Account {
    <#
        .synopsis
        This function will block an office 365 account from signing in
    #>
    [CmdletBinding()]
    Param(    
        [Parameter(Mandatory,HelpMessage='Your Admin Office 365 credentials')]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCredential]$Cred,

        [Parameter(Mandatory,HelpMessage='The user''s office 365 account id')]
        [String]$Email,

        [Parameter(HelpMessage='Set to true if the session is already connected to Azure Active Directory. False by default.')]
        [Boolean]$AlreadyConnected=$false
    )
    
    if(-not $AlreadyConnected) {
        ConnectToMSOnline -Cred $Cred
    }

    Set-MsolUser -UserPrincipalName $Email -BlockCredential $true -ErrorAction Stop
    Write-Host "$Email is now blocked."
}