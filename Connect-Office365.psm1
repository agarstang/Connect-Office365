
function Connect-Office365 {

    param(
    [Parameter(Mandatory=$false)][PSCredential]$UserCredential
    )

    if(-not($UserCredential)) { $UserCredential = Get-Credential }

    Import-Module MSOnline -Scope Global
    Connect-MsolService -Credential $UserCredential

    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-Module (Import-PSSession $ExchangeSession -AllowClobber) -Global 
}


