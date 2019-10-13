# Kill the session to block access to all Office 365 resources
#
# 5 Stages to block and reset the users access to Office 365.
#

# To use this script, first download and setup access to AzureAD, SharePoint Online, Exchange Online

# AzureAD  
# https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-office-365-powershell##connect-with-the-azure-active-directory-powershell-for-graph-module
# SharePoint Online
# https://www.microsoft.com/en-us/download/details.aspx?id=35588

# Exchange Online
# https://www.powershellgallery.com/packages/ExchangeOnlineShell/2.0.3.3

Install-Module -Name AzureAD
Install-Module -Name Microsoft.Online.SharePoint.PowerShell
Install-Module -Name ExchangeOnlineShell

Import-Module AzureAD, ExchangeOnlineShell

# namespace: System.Web.Security
# assembly: System.Web (in System.Web.dll)
# method: GeneratePassword(int length, int numberOfNonAlphanumericCharacters)

# Load "System.Web" assembly in PowerShell console
[Reflection.Assembly]::LoadWithPartialName("System.Web")

# Use to toggle mailbox move
$MoveMailbox = $false 
# Use to toggle CSV input
$UseCSV = $true 
# Tenant admin accoun
$LoginAcctName="tenantadmin@tenant.onmicrosoft.com"
# tenant name for SharePoint connections
$OrgName="tenant" 
# User to disable when not using CSV
$AccountToDisable = "test@tenant.dpmain" 

# Set Login credentials
# If MFA is required, you may need to sign in first to all connections
$LoginCred = Get-Credential

# Connecting to Office 365 endpoints.
# Azure Active Directory
Connect-AzureAD -Credential $LoginCred

# SharePoint Online
Connect-SPOService -Url https://$OrgName-admin.sharepoint.com -Credential $LoginCred

# Exchange Online
Connect-ExchangeOnlineShell -Credential $LoginCred

# Reading in CSV if required. If not, the $AccountToDsiable will be used.
If ($UseCSV)
{
    $userlist = import-Csv -Path .\Users.csv
} else
{
    $userlist = [pscustomobject]@{
                Email = $AccountToDisable
                }
}

foreach ($AccountToDisable in $userlist)
{
    #Calling GeneratePassword Method
    $PW = [System.Web.Security.Membership]::GeneratePassword(16,5)

    Write-Output "Disabling, resetting password for clearing login tokens for $($AccountToDisable.Email) and set password to $PW"

    # Disalbe Account
    Get-AzureADUser -ObjectId $AccountToDisable.Email | Set-AzureADUser -AccountEnabled $false

    # Reset Users Password
    Set-AzureADUserPassword -ObjectId $AccountToDisable.Email -Password (ConvertTo-SecureString -AsPlainText $PW -Force) -ForceChangePasswordNextLogin $true

    # Revoke sessions to SharePoint Online
    Revoke-SPOUserSession -User $AccountToDisable.Email -Confirm:$false

    # Revoke sesion tokens to AzureAD
    Get-AzureADUser -ObjectId $AccountToDisable.Email | Revoke-AzureADUserAllRefreshToken 

    # Force Disconnection on Mailbox by moving mailbox
    if ($MoveMailbox)
    {
        if ( ((get-mailbox -Identity $AccountToDisable.Email -ErrorAction ignore).Database).count -eq 1)
        {
            New-MoveRequest -Identity $AccountToDisable.Email -PrimaryOnly
        }
    }
}
