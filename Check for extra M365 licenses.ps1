#Check if MSOnline module is installed. If it isn't, check for admin rights, escalate if needed, then attempt the install. Once installed the script restarts itself to run again.
if (Get-Module -ListAvailable -Name MSOnline) {
	#Establish connection with M365 module. This will bring up Microsoft's modern auth page (supports MFA). You should login to your partner admin account that has access to the tenants you want to check.
	Connect-MsolService
	#Get list of all M365 tenant IDs
	$customers = Get-MsolPartnerContract -All
	Write-Host "Found $($customers.Count) customers for $((Get-MsolCompanyInformation).displayname)." -ForegroundColor Green
	#Path of the CSV the script creates
	$CSVpath = "$env:USERPROFILE\Desktop\ExtraLicenseReport.csv"
	#Check each customer's license counts, skip counting any free licenses (likely missed some). If additional unused paid licenses are found, add them to the CSV along with the customer's name, license type, and the difference. Microsoft has a page going over which licenses are named what here: https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
	foreach ($customer in $customers) {
		Write-Host "Retrieving license info for $($customer.name)" -ForegroundColor Green
		$licenseCount = Get-MsolAccountSku -TenantId $customer.TenantId  # | Format-List ActiveUnits,ConsumedUnits,AccountSkuId
		#Check each license type for any extras
		foreach ($license in $licenseCount) {
			$licenseType = $license | Select-Object -ExpandProperty AccountSkuId
			if (($licenseType -match "FLOW_FREE") -or ($licenseType -match "POWER_BI_STANDARD") -or ($licenseType -match "TEAMS_EXPLORATORY") -or ($licenseType -match "PROJECT_MADEIRA_PREVIEW_IW_SKU") -or ($licenseType -match "RIGHTSMANAGEMENT_ADHOC") -or ($licenseType -match "FORMS_PRO") -or ($licenseType -match "STREAM") -or ($licenseType -match "SPZA_IW") -or ($licenseType -match "CCIBOTS_PRIVPREV_VIRAL")) {
				Write-Host "$licenseType is a free license, ignoring it." -ForegroundColor Green
			} else {
				$activeUnits = $license | Select-Object -ExpandProperty ActiveUnits
				$consumedUnits = $license | Select-Object -ExpandProperty ConsumedUnits
				$extraUnits = $activeUnits-$consumedUnits
				if ($extraUnits -eq "0"){
					Write-Host "No extra $licenseType licenses found for $($customer.name)" -ForegroundColor Green
				} else {
					Write-Host "Found $extraUnits extra $licenseType for $($customer.name). Adding to CSV..." -ForegroundColor Yellow
					$extraLicenseInfo = [pscustomobject][ordered]@{
						CustomerName	= $customer.Name
						LicenseType		= $licenseType
						ExtraUnits		= $extraUnits
						#TenantId		= $customer.TenantId
					}
					$extraLicenseInfo | Export-CSV -Path $CSVpath -Append -NoTypeInformation
				}
			}
		}
	}
}else{
	Write-Host "MSOnline module is not installed. Attempting to install..."
	#Check if running as admin, if not, restart script as admin to install module.
	if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
		if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
		$CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
		Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
		Exit
		}
	}
	Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
	Install-Module -Name MSOnline -Confirm:$False -Force
	Write-Host "If this install was successful please rerun this script."
}
