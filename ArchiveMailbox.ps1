#This script will convert the specified mailbox to a shared mailbox, hide the account from the GAL, append "- Archived" to the display name, and remove it from all distribution lists.
$choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Y","&N")
while ( $true ) {
	#If Exchange Online module is not installed, check for admin rights, escalate if needed, then attempt to install it. Once installed the script restarts itself to run again.
	if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
		$Admin = Read-Host 'Please enter the Microsoft 365 admin account you would like to use (admin@example.com)'
		$Mailbox = Read-Host 'What email address would you like to archive? (someone@example.com)'
		#Connect to Exchange Online with specified tenant admin
		Connect-ExchangeOnline -UserPrincipalName $admin -ShowProgress $true -ShowBanner:$false
		#Hide specified account from global address list
		Set-Mailbox -Identity $Mailbox -HiddenFromAddressListsEnabled $true
		Write-Host "Account is now hidden from address lists."
		#Get current display name
		$displayname = Get-EXOMailbox -Identity $Mailbox -Properties DisplayName | Select-Object -Property DisplayName | Format-Table -HideTableHeaders | Out-String
		$displayname = $displayname.Trim()
		#Add "- Archived" to end of display name
		Set-User -Identity $Mailbox -DisplayName "$displayname - Archived"
		Write-Host "Account display name updated to: $displayname - Archived"
		$MailboxDN = Get-EXOMailbox $Mailbox
		$DistributionGroups = Get-DistributionGroup
		#Check each distribution group for the account we are archiving. If found, remove the account from the group.
		Write-Host "Attempting to remove account from distribution lists"
		foreach($DistributionGroup in $DistributionGroups){
			$DistributionGroupMembers = Get-DistributionGroupMember -identity $DistributionGroup.Identity | Select-Object -ExpandProperty PrimarySMTPAddress
			foreach ($DistributionGroupMember in $DistributionGroupMembers){
				if ($DistributionGroupMember -eq $MailboxDN.PrimarySMTPAddress){
					try {
						Write-Host "Found $DistributionGroupMember in $($DistributionGroup.name)" -ForegroundColor Yellow
						Remove-DistributionGroupMember $DistributionGroup.Name -Member $DistributionGroupMember -confirm:$false
						Write-Host "Removed $DistributionGroupMember from $($DistributionGroup.name)" -ForegroundColor Green
					}
					catch {
						Write-Host "Failed to remove $($MailboxDN.PrimarySMTPAddress) from $($DistributionGroup.name)" -ForegroundColor Red
					}
				}
			}
		}
		#Convert account to a shared mailbox. This lets us remove the license from the account but keep the emails in it.
		try {
			Set-Mailbox -Identity $Mailbox -Type Shared
			Write-Host "Account has been converted to a shared mailbox. Please remove licenses (if EOP2/Archive please check for litigation hold first.)" -NoNewLine
		}
		catch {
			Write-Host "Failed to convert account to shared mailbox. Please convert manually from the partner/tenant Exchange portal. (if EOP2/Archive please check for litigation hold first.)"
		}
	} 
	else {
		Write-Host "Exchange Online Management module is not installed. Attempting to install..."
		if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
		 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
		  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
		  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
		  Exit
		 }
		}
		Install-Module -Name ExchangeOnlineManagement -Confirm:$False -Force
		Write-Host "If this install was successful please rerun this script."
	}
	#Running the script again allows the user to run this again for another user in the same domain without having to login again.
	$choice = $Host.UI.PromptForChoice("Rerun the script? (y/n)","",$choices,0)
	if ( $choice -ne 0 ) {
		break
	}
}