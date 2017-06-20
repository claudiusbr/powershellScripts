Param(
    [Parameter(Mandatory=$true,HelpMessage='Your Active Directory Admin Credentials')]
    [ValidateNotNullOrEmpty()]
    [PSCredential]$ADCred,

    [Parameter(Mandatory=$true,HelpMessage='The computer name of the machine running AD')]
    [ValidateNotNullOrEmpty()]
    [String]$ServerName,

    [Parameter(Mandatory=$true,HelpMessage='The template user''s SamAccountName')]
    [ValidateNotNullOrEmpty()]
    [String]$TemplateUser

)

$ADSession = New-PSSession -ComputerName $ServerName -Credential $ADCred
Invoke-Command -Session $ADSession -FilePath .\GeneralFunctions.ps1

Describe 'NewAduserFromExisting Integration Test' {
    It 'Creates a new test-user' {
        $script:user = Invoke-Command -Session $ADSession -ScriptBlock {
            Param($TemplateUser)
            NewADUserFromExisting -PreWin2kLogon testunit `
                -FirstName Test -LastName Unit `
                -Email testunit@fake.domain `
                -Domain fake.domain `
                -Password (ConvertTo-SecureString -AsPlainText 'Passworderu234.891!' -Force) `
                -OldUser $TemplateUser -RoleTitle 'Unit Tester' -Organization 'National Test Force' `
                -Department 'Tests for IT'

            ((Get-ADUser testunit -Properties *), (Get-ADUser $TemplateUser -Properties *))
        } -ArgumentList $TemplateUser
    }
    
    It 'Successfully sets up the proxyAddress attribute' {
        $user[0].proxyAddresses | Should Be 'SMTP:testunit@fake.domain'
    }
       
    It "Correctly adds new user to $TemplateUser's security groups" {
        $user[0].MemberOf | Should Be $user[1].MemberOf
    }

    It "Sets up same password change and expiration policy as $TemplateUser's" {
        $user[0].CannotChangePassword | Should Be $user[1].CannotChangePassword
        $user[0].PasswordNeverExpires | Should Be $user[1].PasswordNeverExpires
    }
    
    It 'After-test cleanup' {
        Invoke-Command -Session $ADSession -ScriptBlock {
            Remove-ADUser testunit -Confirm:$false
        }
        Remove-PSSession $ADSession
    }
}