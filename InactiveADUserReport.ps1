#This script will generate 2 CSVs and email them to the specified recipient along with a summary in the body. This should be run through your RMM or a scheduled task on the domain controller. Make sure to update Send-ToEmail with your own account info!

#Ask for admin rights if not already running as an admin
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
	if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
		$CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
		Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
		Exit
	}
}
$RecipientAddress = "ITGUY@yourcompany.com"
function Send-ToEmail([string]$email, [string]$subject, [string]$body, [string]$attachmentpath, [string]$attachmentpath2){
	$message = New-Object Net.Mail.MailMessage;
	$message.From = "sendingaccount@yourcompany.com"; #Account you'd like the emails to come from
	$message.To.Add($email);
	$message.Subject = $subject;
	$message.Body = $body;
	$attachment = New-Object Net.Mail.Attachment($attachmentpath);
	$attachment2 = New-Object Net.Mail.Attachment($attachmentpath2);
	$message.Attachments.Add($attachment);
	$message.Attachments.Add($attachment2);
	
	$smtp = New-Object Net.Mail.SmtpClient("smtp.office365.com", "587"); #Email server address/port
	$smtp.EnableSSL = $true;
	$smtp.Credentials = New-Object System.Net.NetworkCredential("sendingaccount@yourcompany.com", "password"); #Credentials for email account
	$smtp.send($message);
	Write-Host "Email sent";
	$attachment.Dispose();
	$attachment2.Dispose();
}
$ADDomain = (Get-WmiObject -Namespace root\cimv2 -Class Win32_ComputerSystem | Select-Object -ExpandProperty Domain)
$ComputerName = [System.Net.Dns]::GetHostName()+"."+$ADDomain
Write-Host "Found AD domain $ADDomain on $ComputerName"
#Get list of all users that are enabled and have logged in within the last 30 days
$ActiveUsers = Get-Aduser -Filter * -Properties * | Select name,SamAccountName,PasswordExpired,PasswordLastSet,LastLogonDate,Enabled,whenCreated,DistinguishedName,DisplayName,GivenName,SurName | Where-Object Enabled -eq True | Where-Object LastLogonDate -ge $(Get-Date).AddDays(-30)
$ActiveUsers | Export-CSV C:\Windows\Temp\$ADDomain-Active-$(Get-Date -Format "yyyy-MM-dd").csv
#Get list of all users that are enabled but have NOT logged in within the last 30 days
$InactiveUsers = Get-Aduser -Filter * -Properties * | Select name,SamAccountName,PasswordExpired,PasswordLastSet,LastLogonDate,Enabled,whenCreated,DistinguishedName,DisplayName,GivenName,SurName | Where-Object Enabled -eq True | Where-Object SamAccountName -notcontains "NetAdmin" | Where-Object SamAccountName -notcontains "Administrator" | Where-Object SamAccountName -notcontains "Admin" | Where-Object LastLogonDate -le $(Get-Date).AddDays(-30)
$InactiveUsers | Export-CSV C:\Windows\Temp\$ADDomain-Inactive-$(Get-Date -Format "yyyy-MM-dd").csv
Start-Sleep 2
#Generate list of inactive users for email body
$InactiveUserCount = 0
foreach ($User in $InactiveUsers) {
	$InactiveUserCount = $InactiveUserCount+1
	if ($($User.LastLogonDate) -eq $NULL) {
		#Write-Host "$($User.SamAccountName) has never logged in, account created $($User.whenCreated)"
		$InactiveUserList = $InactiveUserList+"`n"+"$($User.SamAccountName) has never logged in, account created $($User.whenCreated)" | Out-String
	}
	else {
		#Write-Host "$($User.SamAccountName) last logged in $($User.LastLogonDate)"
		$InactiveUserList = $InactiveUserList+"`n"+"$($User.SamAccountName) last logged in $($User.LastLogonDate)" | Out-String
	}
}
Write-Host "Found $InactiveUserCount inactive users"
#Send email with both CSVs attached and body showing which accounts are inactive
$Delay = Get-Random -Minimum 1 -Maximum 120 #Random delay up to 2 minutes to avoid rate limiting from sending too many emails at once
Write-Host "Waiting for $Delay seconds before sending email"
Start-Sleep $Delay
Send-ToEmail -email "$RecipientAddress" -subject "AD user report for $ADDomain $(Get-Date -Format MM-dd-yyyy)" -body "Attached is the AD user report for $ADDomain from $ComputerName. Here is the current list of inactive accounts: `n$InactiveUserList `nTotal of $InactiveUserCount inactive accounts found." -AttachmentPath C:\Temp\$ADDomain-Active-$(Get-Date -Format "yyyy-MM-dd").csv -AttachmentPath2 C:\Temp\$ADDomain-Inactive-$(Get-Date -Format "yyyy-MM-dd").csv
#Delete CSVs from C:\Windows\Temp
Remove-Item -Path "C:\Windows\Temp\$ADDomain-Inactive-$(Get-Date -Format "yyyy-MM-dd").csv"
Remove-Item -Path "C:\Windows\Temp\$ADDomain-Active-$(Get-Date -Format "yyyy-MM-dd").csv"