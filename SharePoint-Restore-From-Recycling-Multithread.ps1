#When running this, the time and date you select are your PCs local time and are automatically converted to UTC to avoid timezone issues.
#Thread count is set at line 165 after the ForEach-Object loop.

#WorkingDir and Write-Log are set again in the ForEach-Object loop. If you change this, update it there too.
$WorkingDir = "C:\Temp"
function Write-Log {
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[string]$LogMessage
	)
	Write-Host "$LogMessage"
	Add-Content -Path "$WorkingDir\SharePoint-Restore.log" -Value (Write-Output ("{0} - {1}" -f (Get-Date), $LogMessage))
}

if ($PSVersionTable.PSVersion.Major -ge 7) {
	$choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Y","&N")
	while ($true) {
		#Create calendar objects for date & time selection
		Add-Type -AssemblyName System.Windows.Forms

		# Main Form
		$mainForm = New-Object System.Windows.Forms.Form
		#$font = New-Object System.Drawing.Font("Consolas", 13)
		$mainForm.Text = "Select timeframe"
		$mainForm.Font = $font
		#$mainForm.ForeColor = "White"
		#$mainForm.BackColor = "DarkOliveGreen"
		$mainForm.Width = 300
		$mainForm.Height = 200

		# Start date DatePicker Label
		$StartDatePickerLabel = New-Object System.Windows.Forms.Label
		$StartDatePickerLabel.Text = "Start Date"
		$StartDatePickerLabel.Location = "15, 10"
		$StartDatePickerLabel.Height = 22
		$StartDatePickerLabel.Width = 110
		$mainForm.Controls.Add($StartDatePickerLabel)
		
		# End date DatePicker Label
		$EndDatePickerLabel = New-Object System.Windows.Forms.Label
		$EndDatePickerLabel.Text = "End Date"
		$EndDatePickerLabel.Location = "15, 40"
		$EndDatePickerLabel.Height = 22
		$EndDatePickerLabel.Width = 110
		$mainForm.Controls.Add($EndDatePickerLabel)

		# StartTimePicker Label
		$StartTimePickerLabel = New-Object System.Windows.Forms.Label
		$StartTimePickerLabel.Text = "Start time"
		$StartTimePickerLabel.Location = "15, 70"
		$StartTimePickerLabel.Height = 22
		$StartTimePickerLabel.Width = 110
		$mainForm.Controls.Add($StartTimePickerLabel)
		
		# EndTimePicker Label
		$EndTimePickerLabel = New-Object System.Windows.Forms.Label
		$EndTimePickerLabel.Text = "End time"
		$EndTimePickerLabel.Location = "15, 105"
		$EndTimePickerLabel.Height = 22
		$EndTimePickerLabel.Width = 110
		$mainForm.Controls.Add($EndTimePickerLabel)

		# StartDatePicker
		$StartDatePicker = New-Object System.Windows.Forms.DateTimePicker
		$StartDatePicker.Location = "130, 7"
		$StartDatePicker.Width = "150"
		$StartDatePicker.Format = [windows.forms.datetimepickerFormat]::custom
		$StartDatePicker.CustomFormat = "MM/dd/yyyy"
		$mainForm.Controls.Add($StartDatePicker)
		
		# EndDatePicker
		$EndDatePicker = New-Object System.Windows.Forms.DateTimePicker
		$EndDatePicker.Location = "130, 38"
		$EndDatePicker.Width = "150"
		$EndDatePicker.Format = [windows.forms.datetimepickerFormat]::custom
		$EndDatePicker.CustomFormat = "MM/dd/yyyy"
		$mainForm.Controls.Add($EndDatePicker)

		# StartTimePicker
		$StartTimePicker = New-Object System.Windows.Forms.DateTimePicker
		$StartTimePicker.Location = "130, 70"
		$StartTimePicker.Width = "150"
		$StartTimePicker.Format = [windows.forms.datetimepickerFormat]::custom
		$StartTimePicker.CustomFormat = "HH:mm:ss"
		$StartTimePicker.ShowUpDown = $TRUE
		$mainForm.Controls.Add($StartTimePicker)
		
		# EndTimePicker
		$EndTimePicker = New-Object System.Windows.Forms.DateTimePicker
		$EndTimePicker.Location = "130, 105"
		$EndTimePicker.Width = "150"
		$EndTimePicker.Format = [windows.forms.datetimepickerFormat]::custom
		$EndTimePicker.CustomFormat = "HH:mm:ss"
		$EndTimePicker.ShowUpDown = $TRUE
		$mainForm.Controls.Add($EndTimePicker)


		# OD Button
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Location = "15, 130"
		$okButton.ForeColor = "Black"
		$okButton.BackColor = "White"
		$okButton.Text = "OK"
		$okButton.add_Click({$mainForm.close()})
		$mainForm.Controls.Add($okButton)
		
		if (Get-Module -ListAvailable -Name PnP.PowerShell) {
			Write-Host "PnP PowerShell module found, attempting to connect to 365."
			$SiteURL = Read-Host "Enter the SharePoint site URL (ex. https://reedevelopment.sharepoint.com/sites/Shared)"
			Try {
				Connect-PnPOnline "$SiteURL" -Interactive
				Write-Host "Successfully connected to 365 Cloud."
			}
			Catch {
				Write-Error "Failed to connect to 365 using provided URL/credentials."
			}
			#Get desired timeframe to restore and start restoring files.
			[void] $mainForm.ShowDialog()

			$StartTime = $StartTimePicker | Select -ExpandProperty Value
			$StartDate = $StartDatePicker | Select -ExpandProperty Value
			$EndTime = $EndTimePicker | Select -ExpandProperty Value
			$EndDate = $EndDatePicker | Select -ExpandProperty Value
			
			#Make the dates actually work with PowerShell
			$FinalStart = $(Get-Date -Year $StartDate.Year -Month $StartDate.Month -Day $StartDate.Day -Hour $StartTime.Hour -Minute $StartTime.Minute -Second $StartTime.Second).ToUniversalTime()
			$FinalEnd = $(Get-Date -Year $EndDate.Year -Month $EndDate.Month -Day $EndDate.Day -Hour $EndTime.Hour -Minute $EndTime.Minute -Second $EndTime.Second).ToUniversalTime()
			
			#Grab files that meet the timeframe
			Write-Log "Getting list of files in recycle bin that meet the timeframe."
			$FilesToRestore = Get-PnPRecycleBinItem -FirstStage | ? {$_.DeletedDate -gt $FinalStart -and $_.DeletedDate -lt $FinalEnd}
			#Reverse the list to start from bottom instead
			#[array]::Reverse($FilesToRestore)
			
			#Spit out the list of files and a total count of how many. Ask if the user is sure.
			Write-Host "Files found that meet the timeframe:"
			$FilesToRestore
			Write-Log "Found $($FilesToRestore.count) files to restore."
			Write-Host -nonewline "Files to be restored: $($FilesToRestore.count). Continue? (Y or N): "
			$Response = read-host
			if ($response -ne "Y") { exit }
			
			#Actually restore the items. Thanks BingGPT for the multithreading.
			$FilesToRestore | ForEach-Object -Parallel {
				#Need to declare this function again since ForEach-Object -Parallel runs as separate instances and can't access parent script functions.
				#I tried the $using thing but PowerShell just complains about invalid chars.
				$WorkingDir = "C:\Temp"
				function Write-Log {
					[CmdletBinding()]
					Param (
						[Parameter(Mandatory=$true, Position=0)]
						[string]$LogMessage
					)
					Write-Host "$LogMessage"
					Add-Content -Path "$WorkingDir\SharePoint-Restore.log" -Value (Write-Output ("{0} - {1}" -f (Get-Date), $LogMessage))
				}
				Try {
					$_ | Restore-PnpRecycleBinItem -Force
					Write-Log "Successfully restored $($_.LeafName)"
				}
				Catch {
					Write-Log "Failed to restore $($_.LeafName)"
				}
			} -ThrottleLimit 4 #Max number of restore instances to run at a time. I've found 4 threads avoids ratelimiting from Microsoft. You can try higher if you want.
		}
		else {
			Write-Host "Missing PnP.PowerShell module. Attempting install..."
			if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
				if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
					$CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
					Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
					Exit
				}
			}
			Install-Module -Name PnP.PowerShell -Confirm:$False -Force
			Write-Host "If this install was successful please rerun this script."
		}
		#Running the script again allows the user to run this again for another user in the same domain without having to login again.
		$choice = $Host.UI.PromptForChoice("Rerun the script? (y/n)","",$choices,0)
		if ( $choice -ne 0 ) {
			break
		}
	}
}
else {
	Write-Host "This script requires PowerShell version 7 or higher to run." -ForegroundColor Red
	pause
}
