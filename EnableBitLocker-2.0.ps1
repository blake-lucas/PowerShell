#$Password is used as the boot password for PCs with no TPM.
#If you're using N-Central, set an input parameter on both the automation policy and the "Run PowerShell Script" fields named BitLockerPassword. You should also add an output parameter named BitLockerResult to fill the status field in N-Central.
#If not uncomment the line below and comment the other out.
#$Password = "12InsertPasswordHere53"
$Password = $BitLockerPassword
$WorkingDir = "C:\Temp"
if (($Password -ne $NULL) -and ($Password -ne "") -and ($Password -ne "`n")) {
	Write-Host "Password is not null, converting to SecureString"
	$SecureString = ConvertTo-SecureString $Password -AsPlainText -Force
}
REG ADD HKLM\SOFTWARE\Policies\Microsoft\FVE /v UseAdvancedStartup /t REG_DWORD /d 1 /f
REG ADD HKLM\SOFTWARE\Policies\Microsoft\FVE /v EnableBDEWithNoTPM /t REG_DWORD /d 1 /f
function WaitForEncryption() {
	While ($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty VolumeStatus) -eq "EncryptionInProgress") {
		Write-Log "Waiting for encryption to finish... $(Get-BitLockerVolume | Select-Object -ExpandProperty EncryptionPercentage)%"
		Start-Sleep 1
	}
}
function Write-Log {
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[string]$LogMessage
	)
	Write-Host "$LogMessage"
	Add-Content -Path "$WorkingDir\EnableBitLocker.log" -Value (Write-Output ("{0} - {1}" -f (Get-Date), $LogMessage))
}
#If BitLocker is already enabled with protection turned on, check if TPM is available for use. If TPM is found, remove boot password and add TPM protector instead
if (($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty VolumeStatus) -eq "FullyEncrypted") -and ($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty ProtectionStatus) -eq "On")) {
	Write-Log "C drive is already encrypted and protection is enabled! Double checking if we can use a TPM."
	if ($(Get-TPM | Select-Object -ExpandProperty TpmPresent) -eq "True") {
		Write-Log "TPM Found. Checking if TPM protector already exists."
		if ($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty KeyProtector | Select-Object -ExpandProperty KeyProtectorType) -contains "Tpm") {
			Write-Log "TPM Protector already present! Quitting script."
			$BitLockerResult = "BitLocker enabled and TPM Protector already present!"
			exit
		}
		else {
			Write-Log "No TPM protector found! Adding protector and removing boot password."
			Add-BitLockerKeyProtector -MountPoint "C:" -TpmProtector
			$BLV = Get-BitlockerVolume -MountPoint "C:"
			$PasswordProtector = $BLV.KeyProtector | Where-Object {$PSItem.KeyProtectorType -eq "Password"}
			#Remove-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $PasswordProtector.KeyProtectorId
			$BitLockerResult = "BitLocker already enabled with boot password while TPM is present! Added TPM protector and removed boot password."
			exit
		}
	}
	else {
		Write-Log "No TPM found. Quitting script."
		$BitLockerResult = "BitLocker protection already enabled with boot password and no TPM available."
		exit
	}
}
#If TPM is present and BitLocker is enabled with protection off add TPM protector and turn protection on.
elseif ($(Get-TPM | Select-Object -ExpandProperty TpmPresent) -eq "True") {
	if (($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty VolumeStatus) -eq "FullyEncrypted") -and ($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty ProtectionStatus) -eq "Off") -and ($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty KeyProtector) -eq "")) {
		Write-Log "C is BitLocker encrypted but no key protector was found! Adding TPM protector and enabling protection..."
		Add-BitLockerKeyProtector -MountPoint "C:" -TpmProtector
		Add-BitLockerKeyProtector -MountPoint "C:" -RecoveryPasswordProtector
		Resume-BitLocker -MountPoint "C:"
		$BitLockerResult = "BitLocker already encrypted but no key protectors were found! Added TPM protector and enabled protection."
		#WaitForEncryption
	}
	#Enables protection if protectors already exist but protection is turned off.
	elseif (($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty VolumeStatus) -eq "FullyEncrypted") -and ($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty ProtectionStatus) -eq "Off") -and ($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty KeyProtector | Select-Object -ExpandProperty KeyProtectorType) -contains "Tpm") -and ($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty KeyProtector | Select-Object -ExpandProperty KeyProtectorType) -contains "RecoveryPassword")) {
		Write-Log "TPM and RecoveryPassword protectors found but BitLocker protection is paused! Enabling protection..."
		Resume-BitLocker -MountPoint "C:"
		$BitLockerResult = "TPM and RecoveryPassword protectors already found but protection is paused! Enabled protection again."
	}
	#If no BitLocker stuff was setup previously enable it using TPM.
	else {
		Write-Log "TPM found! Enabling BitLocker with TPM protector."
		#if ($(Get-BitLockerVolume -MountPoint C | Select-Object -ExpandProperty KeyProtector | Select-Object -ExpandProperty KeyProtectorType) -eq "Tpm") {
		#	Write-Log "TPM protector already present! Adding recovery password and encrypting..."
		#	Enable-BitLocker -MountPoint "C:" -EncryptionMethod Aes128 -UsedSpaceOnly -SkipHardwareTest
		#	Start-Sleep 5
		#	Add-BitLockerKeyProtector -MountPoint "C:" -RecoveryPasswordProtector
		#	WaitForEncryption
		#}
		#else {
			Write-Log "TPM protector not found! Adding TPM and recovery password protectors!"
			Enable-BitLocker -MountPoint "C:" -EncryptionMethod Aes128 -UsedSpaceOnly -SkipHardwareTest -TpmProtector
			Start-Sleep 5
			Add-BitLockerKeyProtector -MountPoint "C:" -RecoveryPasswordProtector
			Resume-BitLocker -MountPoint "C:" -ErrorAction SilentlyContinue
			$BitLockerResult = "Enabled BitLocker with TPM protector."
			#WaitForEncryption
		#}
	}
}
#If no TPM is found enable BitLocker using $BitLockerPassword
else {
	Write-Log "No TPM found! Using boot password instead."
	if ($SecureString -eq $NULL) {
		Write-Log "Password was not set! Exiting script!"
		Exit
	}
	Enable-BitLocker -MountPoint "C:" -EncryptionMethod Aes128 -UsedSpaceOnly -SkipHardwareTest -PasswordProtector -Password $SecureString
	Start-Sleep 5
	Add-BitLockerKeyProtector -MountPoint "C:" -RecoveryPasswordProtector
	$BitLockerResult = "Enabled BitLocker using customer's boot password."
	#Add-BitLockerKeyProtector -MountPoint "C:" -Password $SecureString -PasswordProtector
	#WaitForEncryption
}