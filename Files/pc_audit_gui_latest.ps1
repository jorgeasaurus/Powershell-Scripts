<#
    .SYNOPSIS
		Script to generate an HTML Audit report per Endpoint and appends data to CSV 
		Spreadsheet.
    .DESCRIPTION
		The GUI Form takes user input to desginate the correct Endpoint site and Wave 3 
		cutover disposition if already determined among other predetermined data points.
		This data is then appended to a newly generated .htm file and and CSV file or 
		appended to an existing CSV file located on a USB drive in the Jabil Auditor's 
		posession.	
	.EXAMPLE
		- pc_audit_gui_vX.X.ps1 
		- http://bit.ly/jabilaudit
    .OUTPUTS
		Files:
		1. <SITE>_<DATE>_Audit_Report.csv (New Report is generated if either different
		   site is selected or date has changed.)
		2. <DISPO>_<HOSTNAME>_AUDIT_<HHMM>_<DD>_<MM>_<YYYY>.htm (HTM file is generated 
		   and copied to PC and USB.)
		
		Key Data Points:
			'Site Location'
			'Zone'
			'PC No.'
			'Disposition'
			'Computer Asset Tag'
			'First/Last Name'
			'UserName'
			'Model'
			'Peripherals'
			'OS'
			'OS Version'
			'On Domain'
			'Computer IP Address.'
			'Computer MAC Address'
			'USB'
			'Mapped Printer(s)'
			'Printer Model Number'
			'Printer Color Disposition'
			'Share Drives'
			'Specialty Apps'
			'Notes'
			'Eval PC Dispo'
			'ICE PC'

	.NOTES
		Requirements:
			- Windows 7 or Windows 10
			- USB Drive larger than 2GB in Capacity
			- MS Excel 2013 or Later
			- Local Administrative Rights (Preffered but not required)
			- USB Security Not Enabled on Auditor's Endpoint (Preffered but not required)
			- Powershell v2 or Newer

		Changelog:
			Date	Ver		Author	Notes
			-----
			5MAY19	0.1		JS		- Basic console interface for POC.
			6JUN19  1.0     JS      - Bug Fixes, color coding, corrected sites.
		   21JUN19  1.5     JS      - Bug Fixes, added User Interface for Site Audits.
			3JUL19  2.0     JS      - Bug Fixes, added additional fields for input, no
									  longer need to edit CSV file.
			4JUL19  3.0     JS      - UI Changes, added field for 'Considering 
									  Replacement'. Swapped White and Green Button 
									  Locations. 
									  Changed 'TBD' to 'Vendor Managed', added version
									  to Title
			4JUL19	4.0     JS      - UI Changes, Removed field 'Consider Replacement'. 
									  Adjust field placements and sizes.
			4JUL19	4.1     JS      - Removed Copying Files to Local C: Drive
			9JUL19  4.2     JS      - Added Dropdown with Site Specific Zones
			9JUL19  4.3     JS      - Bug Fixes in Drop Downs
		   10JUL19  4.4     JS      - Added 'PC No.' Field.
		   10JUL19  4.5		JS		- Moved 'Peripherals field in CSV next to 'Model'; Added
		   							  'Specialty Apps',ICE PC' and 'Eval PC Dispo'.

#>
	
##*=============================================
##* 				START SCRIPT
##*=============================================

#//////////////////////////////////////////////SETUP/////////////////////////////////////

##*=============================================
##* 		GLOBAL VARIABLE DECLARATION
##*=============================================
$OSVersion = (get-itemproperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"`
			 -Name ProductName).ProductName
$user 	   = whoami
$target    = $env:COMPUTERNAME
$version   = "v4.5 Latest  "

##*=============================================
##* 			POSH v2 COMPATIBILITY
##*=============================================
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

##*=============================================
##* 		DIRECTORY CREATION
##*=============================================

mkdir "$PSScriptRoot\Logs" -Force
mkdir "$PSScriptRoot\Audits" -force

##*=============================================
##* 			ASSEMBLIES
##*=============================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()

# Setup for Transcription Logging
Start-Transcript -Path ("$PSScriptRoot\Logs\$env:computername"+"_"+"$(Get-Date -Format `
yyyyMMddHHmmss).txt") #| Out-Null

#/////////////////////////////////////////////ENDSETUP///////////////////////////////////


#/////////////////////////////////////////////FUNCTIONS//////////////////////////////////
Function confirm() {
	$msgBoxInput =  [System.Windows.MessageBox]::Show("Please confirm site is $site and "+
	"disposition is $dispo.`r`n `r`n Click OK to confirm or Cancel to return to "+
	"previous window.",'Confirm','OKCancel','Warning')
    switch  ($msgBoxInput) {

        'OK' {
			Make_Report
			$Output.Text = "Audit Complete"
			Start-sleep 3
			$Form.close()
			        }
        'Cancel' {
            ## Goes back to main window.
         }
       }
}

function Blue () {
	Write-Host "Do Not Touch Selected, Starting Audit."`n
    $Site = $Sites.SelectedItem.ToString()
    $dispo = $BlueBtn.text
    confirm
}
function Green () {
	Write-Host "Replace Selected, Starting Audit."`n
    $Site = $Sites.SelectedItem.ToString()
    $dispo = $GreenBtn.text
    confirm
}
function Red () {
	Write-Host Convey Selected, Starting Audit.
	#Copy-Item "$PSScriptRoot\poshgui.bat" "C:\JABIL_AUDIT\Domain_Join.bat" -Force
	#Write-Host Domain Join Script Copied to C:\JABIL_AUDIT
	$Site = $Sites.SelectedItem.ToString()
    $dispo = $RedBtn.text
	confirm
	
}
function White () {
	Write-Host Vendor Managed Selected, Starting Audit.`n
    $Site = $Sites.SelectedItem.ToString()
    $dispo = $WhiteBtn.text
    confirm
}

Function Make_Report (){
Function Get-CustomHTML ($Header){
$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($Header)</title>
<META http-equiv=Content-Type content='text/html; charset=windows-1252'>
<meta name="save" content="history">
<style type="text/css">
DIV .expando {DISPLAY: block; FONT-WEIGHT: normal; FONT-SIZE: 8pt; RIGHT: 8px; COLOR: #ffffff; FONT-FAMILY: Arial; POSITION: absolute; TEXT-DECORATION: underline}
TABLE {TABLE-LAYOUT: fixed; FONT-SIZE: 100%; WIDTH: 100%}
*{margin:0}
.dspcont { display:none; BORDER-RIGHT: #B1BABF 1px solid; BORDER-TOP: #B1BABF 1px solid; PADDING-LEFT: 16px; FONT-SIZE: 8pt;MARGIN-BOTTOM: -1px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; WIDTH: 95%; COLOR: #000000; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; BACKGROUND-COLOR: #f9f9f9}
.filler {BORDER-RIGHT: medium none; BORDER-TOP: medium none; DISPLAY: block; BACKGROUND: none transparent scroll repeat 0% 0%; MARGIN-BOTTOM: -1px; FONT: 100%/8px Tahoma; MARGIN-LEFT: 43px; BORDER-LEFT: medium none; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: medium none; POSITION: relative}
.save{behavior:url(#default#savehistory);}
.dspcont1{ display:none}
a.dsphead0 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #FFFFFF; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #CC0000}
a.dsphead1 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #7BA7C7}
a.dsphead2 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #7BA7C7}
a.dsphead1 span.dspchar{font-family:monospace;font-weight:normal;}
td {VERTICAL-ALIGN: TOP; FONT-FAMILY: Tahoma}
th {VERTICAL-ALIGN: TOP; COLOR: #CC0000; TEXT-ALIGN: left}
BODY {margin-left: 4pt} 
BODY {margin-right: 4pt} 
BODY {margin-top: 6pt} 
</style>
<script type="text/javascript">
function dsp(loc){
   if(document.getElementById){
      var foc=loc.firstChild;
      foc=loc.firstChild.innerHTML?
         loc.firstChild:
         loc.firstChild.nextSibling;
      foc.innerHTML=foc.innerHTML=='hide'?'show':'hide';
      foc=loc.parentNode.nextSibling.style?
         loc.parentNode.nextSibling:
         loc.parentNode.nextSibling.nextSibling;
      foc.style.display=foc.style.display=='block'?'none':'block';}}  
if(!document.getElementById)
   document.write('<style type="text/css">\n'+'.dspcont{display:block;}\n'+ '</style>');
</script>
</head>
<body>
<b><font face="Arial" size="5">$($Header)</font></b><hr size="8" color="#CC0000">
<font face="Arial" size="1"><b>Version $Version Created by Jorge Suarez  jorge_suarez1@jabil.com</b></font><br>
<font face="Arial" size="1">Report created on $(Get-Date)</font>
<div class="filler"></div>
<div class="filler"></div>
<div class="filler"></div>
<div class="save">
"@
Return $Report
}


Function Get-CustomHeader0 ($Title){
$Report = @"
		<h1><a class="dsphead0">$($Title)</a></h1>
	<div class="filler"></div>
"@
Return $Report
}

Function Get-CustomHeader ($Num, $Title){
$Report = @"
	<h2><a href="javascript:void(0)" class="dsphead$($Num)" onclick="dsp(this)">
	<span class="expando">show</span>$($Title)</a></h2>
	<div class="dspcont">
"@
Return $Report
}

Function Get-CustomHeaderClose{

	$Report = @"
		</DIV>
		<div class="filler"></div>
"@
Return $Report
}

Function Get-CustomHeader0Close{

	$Report = @"
</DIV>
"@
Return $Report
}

Function Get-CustomHTMLClose{

	$Report = @"
</div>
</body>
</html>
"@
Return $Report
}

Function Get-HTMLTable{
	param([array]$Content)
	$HTMLTable = $Content | ConvertTo-Html
	$HTMLTable = $HTMLTable -replace '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">', ""
	$HTMLTable = $HTMLTable -replace '<html xmlns="http://www.w3.org/1999/xhtml">', ""
	$HTMLTable = $HTMLTable -replace '<head>', ""
	$HTMLTable = $HTMLTable -replace '<title>HTML TABLE</title>', ""
	$HTMLTable = $HTMLTable -replace '</head><body>', ""
	$HTMLTable = $HTMLTable -replace '</body></html>', ""
	Return $HTMLTable
}

Function Get-HTMLDetail ($Heading, $Detail){
$Report = @"
<TABLE>
	<tr>
	<th width='25%'><b>$Heading</b></font></th>
	<td width='75%'>$($Detail)</td>
	</tr>
</TABLE>
"@
Return $Report
}


Write-Output "Collating Detail for $Target"
	$ComputerSystem = Get-WmiObject -computername $Target Win32_ComputerSystem
	switch ($ComputerSystem.DomainRole){
		0 { $ComputerRole = "Standalone Workstation" }
		1 { $ComputerRole = "Member Workstation" }
		2 { $ComputerRole = "Standalone Server" }
		3 { $ComputerRole = "Member Server" }
		4 { $ComputerRole = "Domain Controller" }
		5 { $ComputerRole = "Domain Controller" }
		default { $ComputerRole = "Information not available" }
	}
	
	$OperatingSystems = Get-WmiObject -computername $Target Win32_OperatingSystem
	$TimeZone = Get-WmiObject -computername $Target Win32_Timezone
	$Keyboards = Get-WmiObject -computername $Target Win32_Keyboard
	$SchedTasks = Get-WmiObject -computername $Target Win32_ScheduledJob
	$BootINI = $OperatingSystems.SystemDrive + "boot.ini"
	$RecoveryOptions = Get-WmiObject -computername $Target Win32_OSRecoveryConfiguration
    

    
	switch ($ComputerRole){
		"Member Workstation" { $CompType = "Computer Domain"; break }
		"Domain Controller" { $CompType = "Computer Domain"; break }
		"Member Server" { $CompType = "Computer Domain"; break }
		default { $CompType = "Computer Workgroup"; break }
	}


	$LBTime=$OperatingSystems.ConvertToDateTime($OperatingSystems.Lastbootuptime)
	" "
	#Write-Host "Collecting " $Output.Text
    $Output.Text = "..Regional Options"
    $pbrTest.Value = 10

	$ObjKeyboards = Get-WmiObject -ComputerName $Target Win32_Keyboard
	$keyboardmap = @{
	"00000402" = "BG" 
	"00000404" = "CH" 
	"00000405" = "CZ" 
	"00000406" = "DK" 
	"00000407" = "GR" 
	"00000408" = "GK" 
	"00000409" = "US" 
	"0000040A" = "SP" 
	"0000040B" = "SU" 
	"0000040C" = "FR" 
	"0000040E" = "HU" 
	"0000040F" = "IS" 
	"00000410" = "IT" 
	"00000411" = "JP" 
	"00000412" = "KO" 
	"00000413" = "NL" 
	"00000414" = "NO" 
	"00000415" = "PL" 
	"00000416" = "BR" 
	"00000418" = "RO" 
	"00000419" = "RU" 
	"0000041A" = "YU" 
	"0000041B" = "SL" 
	"0000041C" = "US" 
	"0000041D" = "SV" 
	"0000041F" = "TR" 
	"00000422" = "US" 
	"00000423" = "US" 
	"00000424" = "YU" 
	"00000425" = "ET" 
	"00000426" = "US" 
	"00000427" = "US" 
	"00000804" = "CH" 
	"00000809" = "UK" 
	"0000080A" = "LA" 
	"0000080C" = "BE" 
	"00000813" = "BE" 
	"00000816" = "PO" 
	"00000C0C" = "CF" 
	"00000C1A" = "US" 
	"00001009" = "US" 
	"00001809" = "US" 
	"00010402" = "US" 
	"00010405" = "CZ" 
	"00010407" = "GR" 
	"00010408" = "GK" 
	"00010409" = "DV" 
	"0001040A" = "SP" 
	"0001040E" = "HU" 
	"00010410" = "IT" 
	"00010415" = "PL" 
	"00010419" = "RU" 
	"0001041B" = "SL" 
	"0001041F" = "TR" 
	"00010426" = "US" 
	"00010C0C" = "CF" 
	"00010C1A" = "US" 
	"00020408" = "GK" 
	"00020409" = "US" 
	"00030409" = "USL" 
	"00040409" = "USR" 
	"00050408" = "GK"
	"00000807" = "German_Swiss" 
	"0000100C" = "French_Swiss" 
	"00000810" = "Italian_Swiss"
	}
	$keyb = $keyboardmap.$($ObjKeyboards.Layout | Select -First 1)
	if (!$keyb)
	{ $keyb = "Unknown"
    }
	$Zone = $Zone.text
	$PCNo = $Number.text
	$MyReport = Get-CustomHTML "$Site-----$Zone-----PC No.$PCNo-----$Target-----$Dispo-----AUDIT"
	#$MyReport = Get-CustomHTML "$Target-----$Dispo-----$Site-----$Zone-----PC No.$PCNo"
	$MyReport += Get-CustomHeader0  "$Target Details"
	$MyReport += Get-CustomHeader "2" "General"
		$MyReport += Get-HTMLDetail "Computer Name" ($ComputerSystem.Name)
		$MyReport += Get-HTMLDetail "Computer Role" ($ComputerRole)
		$MyReport += Get-HTMLDetail $CompType ($ComputerSystem.Domain)
		$MyReport += Get-HTMLDetail "Operating System" ($OperatingSystems.Caption)
		$MyReport += Get-HTMLDetail "Service Pack" ($OperatingSystems.CSDVersion)
		$MyReport += Get-HTMLDetail "System Root" ($OperatingSystems.SystemDrive)
		$MyReport += Get-HTMLDetail "Manufacturer" ($ComputerSystem.Manufacturer)
		$MyReport += Get-HTMLDetail "Model" ($ComputerSystem.Model)
		$MyReport += Get-HTMLDetail "Number of Processors" ($ComputerSystem.NumberOfProcessors)
		$MyReport += Get-HTMLDetail "Memory" ($ComputerSystem.TotalPhysicalMemory)
		$MyReport += Get-HTMLDetail "Registered User" ($ComputerSystem.PrimaryOwnerName)
		$MyReport += Get-HTMLDetail "Logged On User" ($User)
		$MyReport += Get-HTMLDetail "Registered Organisation" ($OperatingSystems.Organization)
		$MyReport += Get-HTMLDetail "Last System Boot" ($LBTime)
		$MyReport += Get-HTMLDetail "Site" ($Site)
		$MyReport += Get-HTMLDetail "Zone" ($Zone)
		$MyReport += Get-HTMLDetail "PC No." ($PCNo)
		$MyReport += Get-HTMLDetail "Specialty Apps" ($SpecApps.text)
		$MyReport += Get-HTMLDetail "Notes" ($Notes.text)
		$MyReport += Get-HTMLDetail "Eval PC Dispo" ($EvalPC.text)
		$MyReport += Get-HTMLDetail "ICE PC" ($IcePC.text)
		$MyReport += Get-CustomHeaderClose
		" "
		Write-Host "Collecting " $Output.Text`n
        $Output.Text = "..Logical Disks"
        $pbrTest.Value = 20
		$Disks = Get-WmiObject -ComputerName $Target Win32_LogicalDisk
		$MyReport += Get-CustomHeader "2" "Logical Disk Configuration"
			$LogicalDrives = @()
			Foreach ($LDrive in ($Disks | Where {$_.DriveType -eq 3})){
				$Details = "" | Select "Drive Letter", Label, "File System", "Disk Size (MB)", "Disk Free Space", "% Free Space"
				$Details."Drive Letter" = $LDrive.DeviceID
				$Details.Label = $LDrive.VolumeName
				$Details."File System" = $LDrive.FileSystem
				$Details."Disk Size (MB)" = [math]::round(($LDrive.size / 1MB))
				$Details."Disk Free Space" = [math]::round(($LDrive.FreeSpace / 1MB))
				$Details."% Free Space" = [Math]::Round(($LDrive.FreeSpace /1MB) / ($LDrive.Size / 1MB) * 100)
				$LogicalDrives += $Details
			}
			$MyReport += Get-HTMLTable ($LogicalDrives)
		$MyReport += Get-CustomHeaderClose
		" "
		Write-Host "Collecting " $Output.Text`n
        $Output.Text = "..Network Configuration"
        $pbrTest.Value = 30
		$Adapters = Get-WmiObject -ComputerName $Target Win32_NetworkAdapterConfiguration
		$MyReport += Get-CustomHeader "2" "NIC Configuration"
			$IPInfo = @()
			Foreach ($Adapter in ($Adapters | Where {$_.IPEnabled -eq $True})) {
				$Details = "" | Select Description, "Physical address", "IP Address / Subnet Mask", "Default Gateway", "DHCP Enabled", DNS, WINS
				$Details.Description = "$($Adapter.Description)"
				$Details."Physical address" = "$($Adapter.MACaddress)"
				If ($Adapter.IPAddress -ne $Null) {
				$Details."IP Address / Subnet Mask" = "$($Adapter.IPAddress)/$($Adapter.IPSubnet)"
					$Details."Default Gateway" = "$($Adapter.DefaultIPGateway)"
				}
				If ($Adapter.DHCPEnabled -eq "True")	{
					$Details."DHCP Enabled" = "Yes"
				}
				Else {
					$Details."DHCP Enabled" = "No"
				}
				If ($Adapter.DNSServerSearchOrder -ne $Null)	{
					$Details.DNS =  "$($Adapter.DNSServerSearchOrder)"
				}
				$Details.WINS = "$($Adapter.WINSPrimaryServer) $($Adapter.WINSSecondaryServer)"
				$IPInfo += $Details
			}
			$MyReport += Get-HTMLTable ($IPInfo)
		$MyReport += Get-CustomHeaderClose
		If ((get-wmiobject -ComputerName $Target -namespace "root/cimv2" -list) | Where-Object {$_.name -match "Win32_Product"})
		{
			$Output.Text = "..Software"
			#Updates Progress Bar Value
            $pbrTest.Value = 60
			$MyReport += Get-CustomHeader "2" "Software"
				$MyReport += Get-HTMLTable (get-wmiobject -ComputerName $Target Win32_Product | select Name,Version,Vendor,InstallDate | Sort-Object -Property Name )
			$MyReport += Get-CustomHeaderClose
		}
		Else {
			Write-Output "..Software WMI class not installed"
        }
		" "
		Write-Host "Collecting " $Output.Text`n
		$Output.Text = "..Mapped Drives"
        $pbrTest.Value = 70
		$Shares = Get-wmiobject -ComputerName $Target Win32_MappedLogicalDisk
		$MyReport += Get-CustomHeader "2" "Mapped Drives"
			$MyReport += Get-HTMLTable ($Shares | Select Name, ProviderName)
			#For CSV FILE-----#
			$combined = $Shares | ForEach-Object { $_.Name + ',' + $_.ProviderName }
			$resultshares = $combined -join ')('
			#-----#
			#For Mapped Drive Remapping Batch FILE----#
			#Add-Content "C:\JABIL_AUDIT\map_drives_restore.bat" "@echo off";
			#Get-wmiobject -ComputerName $Target Win32_MappedLogicalDisk  | ForEach-Object {
			#"@net use " +  $_.Name + " " + $_.ProviderName + " path /persistent:yes"} | Out-File "C:\JABIL_AUDIT\map_drives_restore.bat" -Append -Encoding utf8;
			#Add-Content "C:\JABIL_AUDIT\map_drives_restore.bat" ":exit"
			#Add-Content "C:\JABIL_AUDIT\map_drives_restore.bat" "@pause"
			#-----#
		" "
		Write-Host "Collecting " $Output.Text`n

		$MyReport += Get-CustomHeaderClose
        $Output.Text = "..Printers"
        $pbrTest.Value = 80
		$InstalledPrinters =  Get-WmiObject -ComputerName $Target Win32_Printer
		$MyReport += Get-CustomHeader "2" "Printers"
			$MyReport += Get-HTMLTable ($InstalledPrinters | Select Name, Location)
		$MyReport += Get-CustomHeaderClose
		
		$MyReport += Get-CustomHeader "2" "Regional Settings"
			$MyReport += Get-HTMLDetail "Time Zone" ($TimeZone.Description)
			$MyReport += Get-HTMLDetail "Country Code" ($OperatingSystems.Countrycode)
			$MyReport += Get-HTMLDetail "Locale" ($OperatingSystems.Locale)
			$MyReport += Get-HTMLDetail "Operating System Language" (Get-UICulture | Select-Object -ExpandProperty DisplayName)
			$MyReport += Get-HTMLDetail "Keyboard Layout" ($keyb)
		$MyReport += Get-CustomHeaderClose
		
	$MyReport += Get-CustomHeader0Close
	$MyReport += Get-CustomHTMLClose
	$MyReport += Get-CustomHTMLClose

	mkdir "$PSScriptRoot\Audits" -force
	$Filename = "$Site" + "--" + "$Zone" + "--PC No." + "$PCNo" + "--" + "$Target" + "--" + "$Dispo" + "--AUDIT--" +"$(Get-Date -Format yyyyMMdd).htm"
	$Filename = $Filename.ToUpper()
	$MyReport | out-file -encoding ASCII -filepath ("$PSScriptRoot\Audits\$Filename")
	$Output.Text = "Audit saved."
	Write-host $Output.Text`n
	#Updates Progress Bar Value
		$pbrTest.Value = 90
		$i = 99
		While ($i -le 100) {
			$pbrTest.Value = $i
			Start-Sleep -s 1
			$i
			$i += 1
		}
	$Form.Refresh()

	$Output.Text = "Copying to USB..."
	Write-host $Output.Text`n
	Write-host HTML Report saved to "$PSScriptRoot\Audits\$Filename"`n

	# Creates CSV File if one with given name is not found in Directory.
	$CSVReport = "$Site"+"_$(Get-Date -Format yyyyMMdd)"+"_Audit_Report.csv"
	$ScriptPath = "$PSScriptRoot\$CSVReport"
	if (Test-Path -Path $ScriptPath) {
		$rows = [Object[]] (Import-Csv $ScriptPath)
		}
		$addRows =  New-Object PSObject 
		$addRows | add-member Noteproperty    'Site Location'     										$Site
		$addRows | add-member Noteproperty    'Zone'									       			$Zone
		$addRows | add-member Noteproperty	  'PC No.'													$PCNo
		$addRows | add-member Noteproperty    'Disposition'       										$dispo
#		$addRows | add-member Noteproperty    'Consider for Replacement' 								$Consider.text ### Removed in v4
		$addRows | add-member Noteproperty    'Computer Asset Tag'										(Get-WmiObject Win32_ComputerSystem | Select-Object -ExpandProperty Name) 
		$addRows | add-member Noteproperty    'First/Last Name'											$Name.text 
		$addRows | add-member Noteproperty    'UserName'     		    								$user
		$addRows | add-member Noteproperty    'Model'         	    									(Get-WmiObject Win32_ComputerSystem | Select-Object -ExpandProperty Model)
		$addRows | add-member Noteproperty    'Peripherals' 											$Peripherals.text
		$addRows | add-member Noteproperty    'OS'            	    									$OSVersion
		$addRows | add-member Noteproperty	  'OS Version'   		    								(Get-WmiObject Win32_OperatingSystem | Select-Object -ExpandProperty Version)
		$addRows | add-member Noteproperty    'On Domain'     	    									(Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain
		$addRows | add-member Noteproperty    'Computer IP Addresses'    		    					(($Adapter.IPAddress | Select -First 8) -join ",") | Out-String
		$addRows | add-member Noteproperty	  'Computer MAC Address'  									$Adapter.MACaddress
		$addRows | add-member Noteproperty    'USB' 											        $USB.text
		$addRows | add-member Noteproperty	  'Mapped Printer(s)'     									(($InstalledPrinters | Select-Object -ExpandProperty Name) -join ', ')
		$addRows | add-member Noteproperty    'Printer Model Number'							        $PrintModel.text
		$addRows | add-member Noteproperty    'Printer Color Disposition'						        $PrinterDispo.text
		$addRows | add-member Noteproperty	  'Share Drives'										    $resultshares
		$addRows | add-member Noteproperty	  'Specialty Apps'										    $SpecApps.text
		$addRows | add-member Noteproperty	  'Notes' 								                   	$Notes.text
		$addRows | add-member Noteproperty	  'Eval PC Dispo'										    $EvalPC.text
		$addRows | add-member Noteproperty	  'ICE PC'												    $IcePC.text


	$rows + $addRows | Export-Csv "$ScriptPath" -NoTypeInformation
	Write-host "$CSVReport" Generated.
	Write-host CSV saved to "$ScriptPath"


	#Open CSV File
	#Invoke-Item $ScriptPath
	Stop-Transcript
	Write-host All-done.

}
#//////////////////////////////////////////////////ENDFUNCTIONS/////////////////////////////////////////////////////////////////

##*=============================================
##* WINDOWS FORM GUI
##*=============================================

# Init Form
$Form                             = New-Object system.Windows.Forms.Form
$Form.ClientSize                  = '400,700'
$Form.text                        = "JABIL AUDIT TOOL $version"
#$Form.BackColor                   = "#b5adad"
$Form.BackColor                   = "#ffffff"
$Form.TopMost                     = $false
$Form.StartPosition               = "CenterScreen"
$Form.FormBorderStyle             = 'Fixed3D'
$Form.MaximizeBox                 = $false
#$Form.ControlBox                  = $false

$iconBase64      = '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UYAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAQQQAAAAcRAAAAMBBBAAAABxEAAAAwEEEAAAAHEQAAADAOdL/AGjsLcM8AdqLAAAAAElFTkSuQmCC'
$iconBytes       = [Convert]::FromBase64String($iconBase64)
$stream          = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$iconImage       = [System.Drawing.Image]::FromStream($stream, $true)
$Form.Icon       = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())


# Init Blue Button
$BlueBtn                          = New-Object system.Windows.Forms.Button
$BlueBtn.BackColor                = "#0957b1"
$BlueBtn.text                     = "Do Not Touch"
$BlueBtn.width                    = 92
$BlueBtn.height                   = 58
$BlueBtn.location                 = New-Object System.Drawing.Point(7,640)
$BlueBtn.Font                     = 'Microsoft Sans Serif,10,style=Bold'

# Init Green Button
$GreenBtn                         = New-Object system.Windows.Forms.Button
$GreenBtn.BackColor               = "#28b608"
$GreenBtn.text                    = "Replace"
$GreenBtn.width                   = 92
$GreenBtn.height                  = 58
$GreenBtn.location                = New-Object System.Drawing.Point(300,640)
$GreenBtn.Font                    = 'Microsoft Sans Serif,10,style=Bold'

# Init Red Button
$RedBtn                           = New-Object system.Windows.Forms.Button
$RedBtn.BackColor                 = "#d80a0a"
$RedBtn.text                      = "Convey"
$RedBtn.width                     = 92
$RedBtn.height                    = 58
$RedBtn.location                  = New-Object System.Drawing.Point(202,640)
$RedBtn.Font                      = 'Microsoft Sans Serif,10,style=Bold'

# Init White Button
$WhiteBtn                         = New-Object system.Windows.Forms.Button
$WhiteBtn.BackColor               = "#b5adad"
#$WhiteBtn.BackColor               = "#ffffff"
$WhiteBtn.text                    = "Vendor Managed"
$WhiteBtn.width                   = 92
$WhiteBtn.height                  = 58
$WhiteBtn.location                = New-Object System.Drawing.Point(104,640)
$WhiteBtn.Font                    = 'Microsoft Sans Serif,10,style=Bold'

$Image 							  = [system.drawing.image]::FromFile("$PSScriptroot\265_main.jpg")
$pictureBox 					  = new-object Windows.Forms.PictureBox 
$pictureBox.width				  = 327
$pictureBox.height   			  = 180
$pictureBox.location              = New-Object System.Drawing.Point(33,9)
$pictureBox.Image				  = $image
$pictureBox.SizeMode              = [System.Windows.Forms.PictureBoxSizeMode]::zoom

$PCNAME                           = New-Object system.Windows.Forms.TextBox
$PCNAME.multiline                 = $false
$PCNAME.text                      = "$env:COMPUTERNAME"
$PCNAME.width                     = 200+33
$PCNAME.height                    = 123
$PCNAME.enabled                   = $false
$PCNAME.location                  = New-Object System.Drawing.Point(123,208)
$PCNAME.Font                      = 'Microsoft Sans Serif,10'

$Label2                           = New-Object system.Windows.Forms.Label
$Label2.text                      = "PC Name:"
$Label2.AutoSize                  = $true
$Label2.width                     = 200
$Label2.height                    = 108
$Label2.location                  = New-Object System.Drawing.Point(33,212)
$Label2.Font                      = 'Microsoft Sans Serif,10'



$descriptions = @('Select Site','Balsthal','Bettlach','Grenchen','Haegendorf','Le Locle','Mezzovico','Raron')



$Sites_SelectedIndexChanged=
    {
        Switch ($Sites.text)
       {
            
            'Balsthal'
            {
				$envnames = @('Select Zone',
				'Zone 10 (A)',
                'Zone 11 (A)',
                'Zone 20 (A)',
                'Zone 21 (A)',
                'Zone 30 (A)',
                'Zone 31 (A)',
                'Zone 40 (B)',
                'Zone 41 (B)',
                'Zone 50 (B)',
                'Zone 51 (B)',
                'Zone 52 (B)',
                'Zone 53 (B)'
                )
            }
            'Bettlach'
            {
				$envnames = @('Select Zone',
				'Zone 11',
                'Zone 12',
                'Zone 13',
                'Zone 14',
                'Zone 20',
                'Zone 21',
                'Zone 22',
                'Zone 23',
                'Zone 30',
                'Zone 31',
                'Zone 32',
                'Zone 33'
                )
            }

            'Grenchen' 
            {
				$envnames = @('Select Zone',
				'Zone 10',
                'Zone 11',
                'Zone 12',
                'Zone 13',
                'Zone 14',
                'Zone 20',
                'Zone 21',
                'Zone 22',
                'Zone 23',
                'Zone 24',
                'Zone 25',
                'Zone 30',
                'Zone 31',
                'Zone 32',
                'Zone 33',
                'Zone 34',
                'Zone 35'
                )
            }
            'Haegendorf'
            {
				$envnames = @('Select Zone',
				'Zone 10',
                'Zone 11',
                'Zone 12',
                'Zone 20',
                'Zone 21',
                'Zone 22',
                'Zone 23',
                'Zone 30',
                'Zone 31',
                'Zone 32'
                )
            }
            'Le Locle'
            {
				$envnames = @('Select Zone',
				'CB1 Zone 10',
                'CB1 Zone 11',
                'CB1 Zone 12',
                'CB1 Zone 13',
                'CB1 Zone 14',
                'CB1 Zone 20',
                'CB1 Zone 21',
                'CB1 Zone 22',
                'CB1 Zone 30',
                'CB1 Zone 31',
                'CB1 Zone 32',
                'CB2 Zone 40',
                'CB2 Zone 41',
                'CB2 Zone 42',
                'CB2 Zone 43',
                'CB2 Zone 50',
                'CB2 Zone 51',
                'CB2 Zone 52',
                'CB2 Zone 60',
                'CB2 Zone 61',
                'CB2 Zone 62',
                'CB2 Zone 63',
                'CB2 Zone 70'
                )
            }
            'Mezzovico'
            {
				$envnames = @('Select Zone',
				'Zone 10',
                'Zone 21',
                'Zone 22',
                'Zone 23',
                'Zone 24',
                'Zone 25',
                'Zone 26',
                'Zone 30',
                'Zone 31',
                'Zone 32',
                'Zone 33',
                'Zone 34',
                'Zone 35',
                'Zone 40'
                )
            }
            'Raron'
            {
				$envnames = @('Select Zone',
				'Zone 10',
                'Zone 11',
                'Zone 12',
                'Zone 13',
                'Zone 14',
                'Zone 15',
                'Zone 16',
                'Zone 17',
                'Zone 20',
                'Zone 21',
                'Zone 22'
                )

        }

        }
$Zone.Remove_SelectedIndexChanged($Zone_SelectedIndexChanged)
$Zone.DataSource = $envnames
$Zone.add_SelectedIndexChanged($Zone_SelectedIndexChanged)
    }


$Sites                            = New-Object system.Windows.Forms.ComboBox
@('Balsthal','Bettlach','Grenchen','Haegendorf','Le Locle','Mezzovico','Raron') | ForEach-Object {[void] $Sites.Items.Add($_)}
$Sites.width                      = 200+33
$Sites.height                     = 123
$Sites.location                   = New-Object System.Drawing.Point(123,237)
$Sites.Font                       = 'Microsoft Sans Serif,10'
$Sites.SelectedIndex              = 0
$Sites.DataSource = $descriptions
$Sites.add_SelectedIndexChanged($Sites_SelectedIndexChanged)


$SiteLabel                        = New-Object system.Windows.Forms.Label
$SiteLabel.text                   = "Site: "
$SiteLabel.AutoSize               = $true
$SiteLabel.width                  = 200
$SiteLabel.height                 = 108
$SiteLabel.location               = New-Object System.Drawing.Point(33,243)
$SiteLabel.Font                   = 'Microsoft Sans Serif,10'

$Zone                            = New-Object system.Windows.Forms.ComboBox
#$Zone.multiline                   = $false
$Zone.text                        = ""
$Zone.width                       = 100
$Zone.height                      = 123
$Zone.enabled                     = $true
$Zone.location                    = New-Object System.Drawing.Point(123,268)
$Zone.Font                        = 'Microsoft Sans Serif,10'
$Zone.add_SelectedIndexChanged($Zone_SelectedIndexChanged)


$Zonelabel                        = New-Object system.Windows.Forms.Label
$Zonelabel.text                   = "Zone:"
$Zonelabel.AutoSize               = $true
$Zonelabel.width                  = 200
$Zonelabel.height                 = 108
$Zonelabel.location               = New-Object System.Drawing.Point(33,271)
$Zonelabel.Font                   = 'Microsoft Sans Serif,10'

$Number                             = New-Object system.Windows.Forms.TextBox
$Number.multiline                   = $false
$Number.text                        = " "
$Number.width                       = 60
$Number.height                      = 123
$Number.enabled                     = $true
$Number.location                    = New-Object System.Drawing.Point(294,268)
$Number.Font                        = 'Microsoft Sans Serif,10'			

$Numberlabel                        = New-Object system.Windows.Forms.Label
$Numberlabel.text                   = "PC No.:"
$Numberlabel.AutoSize               = $true
$Numberlabel.width                  = 200
$Numberlabel.height                 = 108
$Numberlabel.location               = New-Object System.Drawing.Point(230,271)
$Numberlabel.Font                   = 'Microsoft Sans Serif,10'

$Name                             = New-Object system.Windows.Forms.TextBox
$Name.multiline                   = $false
$Name.text                        = " "
$Name.width                       = 200+33
$Name.height                      = 123
$Name.enabled                     = $true
$Name.location                    = New-Object System.Drawing.Point(123,299)
$Name.Font                        = 'Microsoft Sans Serif,10'

$Namelabel                        = New-Object system.Windows.Forms.Label
$Namelabel.text                   = "Name:"
$Namelabel.AutoSize               = $true
$Namelabel.width                  = 200
$Namelabel.height                 = 108
$Namelabel.location               = New-Object System.Drawing.Point(33,302)
$Namelabel.Font                   = 'Microsoft Sans Serif,10'

$Peripherals                      = New-Object system.Windows.Forms.TextBox
$Peripherals.multiline            = $false
$Peripherals.text                 = " "
$Peripherals.width                = 200+33
$Peripherals.height               = 61
$Peripherals.enabled              = $true
$Peripherals.location             = New-Object System.Drawing.Point(123,330)
$Peripherals.Font                 = 'Microsoft Sans Serif,10'

$Peripheralslabel                 = New-Object system.Windows.Forms.Label
$Peripheralslabel.text            = "Peripherals:"
$Peripheralslabel.AutoSize        = $true
$Peripheralslabel.width           = 200
$Peripheralslabel.height          = 108
$Peripheralslabel.location        = New-Object System.Drawing.Point(33,333)
$Peripheralslabel.Font            = 'Microsoft Sans Serif,10'

$USB                              = New-Object system.Windows.Forms.TextBox
$USB.multiline                    = $false
$USB.text                         = " "
$USB.width                        = 200+33
$USB.height                       = 61
$USB.enabled                      = $true
$USB.location                     = New-Object System.Drawing.Point(123,361)
$USB.Font                         = 'Microsoft Sans Serif,10'

$USBlabel                         = New-Object system.Windows.Forms.Label
$USBlabel.text                    = "USB:"
$USBlabel.AutoSize                = $true
$USBlabel.width                   = 200
$USBlabel.height                  = 108
$USBlabel.location                = New-Object System.Drawing.Point(33,364)
$USBlabel.Font                    = 'Microsoft Sans Serif,10'

$PrintModel                       = New-Object system.Windows.Forms.TextBox
$PrintModel.multiline             = $false
$PrintModel.text                  = " "
$PrintModel.width                 = 200+33
$PrintModel.height                = 61
$PrintModel.enabled               = $true
$PrintModel.location              = New-Object System.Drawing.Point(123,391)
$PrintModel.Font                  = 'Microsoft Sans Serif,10'

$PrintModellabel                  = New-Object system.Windows.Forms.Label
$PrintModellabel.text             = "Printer Model:"
$PrintModellabel.AutoSize         = $true
$PrintModellabel.width            = 200
$PrintModellabel.height           = 108
$PrintModellabel.location         = New-Object System.Drawing.Point(33,394)
$PrintModellabel.Font             = 'Microsoft Sans Serif,10'

$PrinterDispo                     = New-Object system.Windows.Forms.ComboBox
@('None','Convey (Red)','Do Not Touch (Blue)') | ForEach-Object {[void] $PrinterDispo.Items.Add($_)}
$PrinterDispo.width               = 200+33
$PrinterDispo.height              = 123
$PrinterDispo.location            = New-Object System.Drawing.Point(123,421)
$PrinterDispo.Font                = 'Microsoft Sans Serif,10'
$PrinterDispo.SelectedIndex       = 0

$PrinterDispoLabel                = New-Object system.Windows.Forms.Label
$PrinterDispoLabel.text           = "Printer Dispo: "
$PrinterDispoLabel.AutoSize       = $true
$PrinterDispoLabel.width          = 200
$PrinterDispoLabel.height         = 108
$PrinterDispoLabel.location       = New-Object System.Drawing.Point(33,423)
$PrinterDispoLabel.Font           = 'Microsoft Sans Serif,10'

$SpecApps                            = New-Object system.Windows.Forms.TextBox
$SpecApps.multiline                  = $false
$SpecApps.text                       = " "
$SpecApps.width                      = 200+33
$SpecApps.height                     = 50
$SpecApps.enabled                    = $true
$SpecApps.location                   = New-Object System.Drawing.Point(123,451)
$SpecApps.Font                       = 'Microsoft Sans Serif,10'

$SpecAppsLabel                       = New-Object system.Windows.Forms.Label
$SpecAppsLabel.text                  = "Specialty`nApplications:"
$SpecAppsLabel.AutoSize              = $true
$SpecAppsLabel.width                 = 50
$SpecAppsLabel.height                = 150
$SpecAppsLabel.location              = New-Object System.Drawing.Point(33,445)
$SpecAppsLabel.Font                  = 'Microsoft Sans Serif,10'

$EvalPC                            = New-Object system.Windows.Forms.ComboBox
@('Yes','No') | ForEach-Object {[void] $EvalPC.Items.Add($_)}
$EvalPC.text                        = ""
$EvalPC.width                       = 100
$EvalPC.height                      = 123
$EvalPC.enabled                     = $true
$EvalPC.location                    = New-Object System.Drawing.Point(123,480)
$EvalPC.Font                        = 'Microsoft Sans Serif,10'
$EvalPC.SelectedIndex               = 1

$EvalPClabel                        = New-Object system.Windows.Forms.Label
$EvalPClabel.text                   = "Eval PC Dispo:"
$EvalPClabel.AutoSize               = $true
$EvalPClabel.width                  = 200
$EvalPClabel.height                 = 108
$EvalPClabel.location               = New-Object System.Drawing.Point(33,485)
$EvalPClabel.Font                   = 'Microsoft Sans Serif,10'

$IcePC                             = New-Object system.Windows.Forms.ComboBox
@('Yes','No') | ForEach-Object {[void] $IcePC.Items.Add($_)}
$IcePC.text                        = " "
$IcePC.width                       = 60
$IcePC.height                      = 123
$IcePC.enabled                     = $true
$IcePC.location                    = New-Object System.Drawing.Point(294,480)
$IcePC.Font                        = 'Microsoft Sans Serif,10'	
$IcePC.SelectedIndex              = 1		

$IcePClabel                        = New-Object system.Windows.Forms.Label
$IcePClabel.text                   = "ICE PC:"
$IcePClabel.AutoSize               = $true
$IcePClabel.width                  = 200
$IcePClabel.height                 = 108
$IcePClabel.location               = New-Object System.Drawing.Point(230,483)
$IcePClabel.Font                   = 'Microsoft Sans Serif,10'

$Notes                            = New-Object system.Windows.Forms.TextBox
$Notes.multiline                  = $true
$Notes.text                       = " "
$Notes.width                      = 200+33
$Notes.height                     = 50
$Notes.enabled                    = $true
$Notes.location                   = New-Object System.Drawing.Point(123,510)
$Notes.Font                       = 'Microsoft Sans Serif,10'

$NotesLabel                       = New-Object system.Windows.Forms.Label
$NotesLabel.text                  = "Notes:"
$NotesLabel.AutoSize              = $true
$NotesLabel.width                 = 50
$NotesLabel.height                = 150
$NotesLabel.location              = New-Object System.Drawing.Point(33,512)
$NotesLabel.Font                  = 'Microsoft Sans Serif,10'

$Output                           = New-Object system.Windows.Forms.Label
$Output.AutoSize                  = $true
$Output.width                     = 220
$Output.height                    = 108
$Output.location                  = New-Object System.Drawing.Point(135,564)
$Output.Font                      = 'Microsoft Sans Serif,10'
$Output.Text                      = "Progress"

# Init ProgressBar
$pbrTest                          = New-Object System.Windows.Forms.ProgressBar
$pbrTest.Maximum		 		  = 100
$pbrTest.Minimum				  = 0
$pbrTest.Location				  = new-object System.Drawing.Size(25,594)
$pbrTest.size					  = new-object System.Drawing.Size(350,30)
$i								  = 0


$Form.controls.AddRange(@(
		$BlueBtn,
		$GreenBtn,
		$RedBtn,
		$WhiteBtn,
		$Jabillogo,
		$Sites,
		$SiteLabel,
		$PCNAME,
		$Label2,
		$Output,
		$pbrTest,
		$pictureBox,
		$Zone,
		$Zonelabel,
		$Name,
		$Namelabel,
		$Peripherals,
		$Peripheralslabel,
		$PhoneModel,
		$PhoneModellabel,
		$USB,
		$USBlabel,
		$PrinterDispo,
		$PrinterDispoLabel,
		$PrintModel,
		$PrintModellabel,
		$Notes,
		$NotesLabel,
		$Numberlabel,
		$Number,
		$SpecApps,
		$SpecAppsLabel,
		$EvalPC,
		$EvalPClabel,
		$IcePC,
		$IcePClabel ))

$BlueBtn.Add_Click({Blue})
$GreenBtn.Add_Click({Green})
$RedBtn.Add_Click({Red})
$WhiteBtn.Add_Click({White})

$Site = $Sites.SelectedIndex




######################################



[void]$Form.ShowDialog()

##*=============================================
##* END SCRIPT
##*=============================================

# SIG # Begin signature block
# MIIFkQYJKoZIhvcNAQcCoIIFgjCCBX4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUGYNqhADbgCuFpLWwwAwkH5pl
# 4ZygggMgMIIDHDCCAgSgAwIBAgIQJ3wEiYQ/ip9CXSfpEB84ADANBgkqhkiG9w0B
# AQsFADAmMSQwIgYDVQQDDBtKb3JnZSdzIFNpZ25pbmcgQ2VydGlmaWNhdGUwHhcN
# MTkwNTMxMTgxOTIzWhcNMjAwNTMxMTgzOTIzWjAmMSQwIgYDVQQDDBtKb3JnZSdz
# IFNpZ25pbmcgQ2VydGlmaWNhdGUwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
# AoIBAQCrRmAw1Y/nYDJ6AjmX713OruONVi8q2YQMc1E/vw5SjwPyHhC+RcMdkgGC
# VGDrZfWkQgqle0ikyjHfQWUMfFtTUHEh0KDC0zDY4yKUV/NFpklYweNpd/VfKWgc
# Jzgpao+uebR2qxQ3Cz7zajQMixC+tw/7hOZ+ykpaKl1wQOlL7wxFEWOW38PUmwki
# XCDdMDrVVB0wmrADWKUsqaL4zuCFWrmoYP/fxjyoIadD72LVTxFHoV9ObWAdrQn4
# l/UlZV0QaKvasLIADy5T+YT5thaDtazJgDjXZSsgVFPG3/22U/AFrIFM3WytGdVy
# ebNEPxY+EiINekk/HV00z5zE7zsFAgMBAAGjRjBEMA4GA1UdDwEB/wQEAwIHgDAT
# BgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQU8i+06pVpPa3aWlk7PPrCC6Jz
# csQwDQYJKoZIhvcNAQELBQADggEBAAtizZPZfLVfhSCJvZPvVuL/BznPTjXm9WrX
# hiQek0tVrAeqJ8gfjT77ZKQbir5LYi2s1KqZreLVtNB+QWzgkfxeAvBCVK48kJJD
# EyREQUtDE/lhyUUaNtdQXQBRYjnxgrNbGLbjEc5V/XXriLlAqPm0VLg0zwBPvFj1
# IVkms/40lApiOkJGvom2Fl5wlTuJzDeLtS6czUSVN8J88vOcKiYNoO5Qu9xXwvGw
# 8bJrnsqTToCHPuHZVa//hDVXYrnILxWP2JOVQntuJ+pv2rkDdGOXE7MmCia3iUo7
# ncsLLXT9mf29jFs6rEF8q5IgpvpL+YgUEtfBQY6dogzKp0yUPqkxggHbMIIB1wIB
# ATA6MCYxJDAiBgNVBAMMG0pvcmdlJ3MgU2lnbmluZyBDZXJ0aWZpY2F0ZQIQJ3wE
# iYQ/ip9CXSfpEB84ADAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAA
# oQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUSjzuoZ67H+Hu2E9jMDp8rJt7
# /nAwDQYJKoZIhvcNAQEBBQAEggEAGS7+16Aj5IzcgZTU4DlXBQvKaXv429sXPhJ5
# CozeNfXJNBRbZ2BPNaiTgPCFQA/ciSHlsjL0/mYxZe+7Ml2qrBkAp+ZHzsj9Y2OT
# hrx5CEE8+VuATMxrw3u85PfOjBNYBsxm7hpY063cPFs8tTKxb1VjJsw/HyXaqgOC
# IR630rP8YqkL2cmHez+YCupaS+J9t6+gdejnNb5ZQYqCm09/w4fqwA//im1yMExL
# APAdTsdYPwPD7bizwKqwJUB+DaVx2tGsEjjbAE5mEpiVkyeArufboSnODGZfmkIP
# 86nu6uV1kkwz1DIIYhKZ3JY+cFht4xq3hPUGYmjq5NvmxYWAFg==
# SIG # End signature block
