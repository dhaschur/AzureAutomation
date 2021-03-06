param ([switch]$SendEmail = $false, [string]$OneOffEmailAddr = "default")
# By default do not send email.  Therefore it can be run from the command line without mailing everyone.

# .\COMPANYExchangeReport-New.ps1     		#Runs without sending any email
# .\COMPANYExchangeReport-New.ps1  -SendEmail  		#Runs and sends mail to the builtin default addresses
# .\COMPANYExchangeReport-New.ps1  -SendEmail  name@domain.com   		#Runs and sends email to the provided address only
# In Scheduled Tasks use -->  powershell -command "& {C:\WINDOWS\Scripts\COMPANYExchangeReport-new.ps1 -SendEmail}"

cls
#$DebugPreference = "SilentlyContinue"  #This is default
$DebugPreference = "Continue"  #This causes more output
#$DebugPreference = "Inquire"  #This causes output and queries to continue or stop
#$DebugPreference = "Stop"  #This causes output and stops

#Load up the Report Library Functions.  ReportLibrary.ps1 must be in the same dir as the report script
Write-Debug "Load Report Libary Functions"
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptDir = Split-Path $MyInvocation.MyCommand.Path
. "$ScriptDir\..\ReportLibrary\ReportLibrary.ps1"

Confirm-ExchangeSnapIn

## Setup Variables ######
$ServerName = $Env:COMPUTERNAME
$FromServer = "&nbsp;<span class=""style5"">" + $ServerName + "</span>"
$now = get-date
$date = get-date -format d #get-date -uformat "%m-%d-%Y"
$TZ = Get-TimeZoneName
$ServerSkipList = "C:\WINDOWS\Scripts\ExServersToSkip.csv"
$Logo = '<p style= "font-family:trebuchet ms;color:#0085c3;font-size:110%;font-weight:Bold">COMPANY NAME<p>&nbsp;'
## Company Info and Mail
$ReportType = "Exchange"
$Company = "COMPANY"
$Subject = $Company + " " + $ReportType + " Report - " + $date
$Header = $Company + " " + $ReportType + " Report - Report as on " + $now + " " + $TZ + $FromServer + "<br>" + "`r`n"
$SMTPHost = "SERVER.COM"
$DefaultFrom = "ExchangeOp@COMPANYNAME.org"
$emailMessage = "Exchange Report"
$DefaultTo = ("USER@COMPANYNAME.org")
$DefaultCc =  ("")
$DefaultBcc =  ("")
$whiteSpaceReportName = "REPORT.csv"
$whiteSpaceReportPath = "C:\REPORTS"

$attach = ""
$filelocation = $Company + $ReportType + "Report.html"
$Code = @()
$Write = @()



## Function WhiteSpace
Function Get-WhiteSpace([string]$server, [string]$MailboxDatabaseName, [array]$WSEvents) {
	$Array = @()
	$Result = @()
	If ($WSEvents -ne $null) {
		#$Check = $WSpace -is [array]
		$WSpace = $WSEvents | Where-Object {$_.Message.Contains($MailboxDatabaseName)}
		If ($WSpace -eq $null) {
			Write-Debug "WSpace is null"
			$Result = "No defrag last day"
		} elseif ($WSpace -is [array]) {
			Write-Debug 'WSpace is Array: ' #($WSpace -is [array])
			ForEach ($item in $WSpace) {
				$Array = $item.message.split("`"")
				$EndString = $Array[2].indexof("megabytes ")-6
				#Write-Debug $mbnamearray[1]
				[int]$Result = $Array[2].Substring(5,$EndString)
			}
		} else {
			Write-Debug 'WSpace is not Array: ' #($WSpace -is [array])
			$Array = $WSpace.message.split("`"")
			$EndString = $Array[2].indexof("megabytes ")-6
			#Write-Debug $mbnamearray[1]
			[int]$Result = $Array[2].Substring(5,$EndString)			 
		}
	} else {
		Write-Debug "WSpace is null"
		$Result = "No defrag last day"
	}
	
	$Result
} #end of Function Get-Whitespace

## HEAD HTML Code
function Head() {
#$Write = ""
$Write += "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01//EN""  ""http://www.w3.org/TR/html4/strict.dtd"">" + "`r`n"
$Write += "<html>" + "`r`n"
$Write += "<head>" + "`r`n"
$Write += "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">" + "`r`n"
$title = "<title>" + $Subject + "</title>" + "`r`n"
$Write += $title
$Head = @'
<style>
BODY{font-family:tahoma;} 
TABLE{border-width: 1px;cellspacing: 0;cellpadding: 3;}
TH{border-width: 1px;font-size: 11px;text-align: center;background-color: #999999;font-weight: bold;}
TD{border-width: 1px;font-size: 9px;text-align: center;border-width: 1px;border-style: solid;}
.style5{color: #FFFFFF;}
.style9{background-color: #000000;}
.style10{font-family: Verdana;font-size: 14px;font-weight: bold;}
.style11{font-size: 9px;vertical-align: text-top;}
.style12{font-size: 9px;color: #C0C0C0;}
.style13{font-family: Tahoma;font-size: 9px;color: #FFFFFF;border: 1px solid #C0C0C0;}
.style14{color: #FE000B;font-family: Verdana;font-size: 14px;font-weight: bold;}
</style>
'@
$Write += $Head
$Write += "</head><body>" + "`r`n"
$Write += "<H6>" + "`r`n"
$Write += $Header
$Write += "<br>" + "`r`n"
$Write += "Backup: <span style=""background-color:yellow"">YELLOW</span> = &gt; 1 day &nbsp; <span style=""background-color:red"">RED</span> = &gt; 2 days <br>" + "`r`n"
$Write += "Whitespace: <span style=""background-color:yellow"">YELLOW</span> = &gt; 20% &nbsp; <span style=""background-color:red"">RED</span> = &gt; 30% <br>" + "`r`n"
$Write += "Free Space: <span style=""background-color:yellow"">YELLOW</span> = &gt; 15% &nbsp; <span style=""background-color:red"">RED</span> = &gt; 5% <br>" + "`r`n"
$Write += "</H6>" + "`r`n"
$Write += "<table>" + "`r`n"
Return $Write
} #end of Function Head

## HEADER HTML Code
function Header() {
#$Write = ""
$Write += "<tr>" + "`r`n"
$Write += "<th>Server\StorageGroup\Database</th>" + "`r`n"
$Write += "<th>Size (GB)</th>" + "`r`n"
$Write += "<th>Mailbox Count</th>" + "`r`n"
$Write += "<th>LastFullBackup</th>" + "`r`n"
$Write += "<th>IncrementalBackup</th>" + "`r`n"
$Write += "<th>WhiteSpace (MB)</th>" + "`r`n"
$Write += "<th>Capacity (GB)</th>" + "`r`n"
$Write += "<th>Free Space (GB)</th>" + "`r`n"
$Write += "<th>Free Space (%)</th>" + "`r`n"
$Write += "</tr>" + "`r`n"
Return $Write
} # end of Function Header

## BODY HTML Code
function Body($objItem,$MDBSize,$MBXCount,$BackupDays,$IncBackupDays,$WhiteSpace,$DiskCapacity,$FreeOnStorage,$FreeOnStoragePct) {
	$Write += "<tr>" + "`r`n"
	$Write += "<td>$objItem</td>" + "`r`n"
	#Write-Debug "Size (GB): " $MDBSize
	$Write += "<td>$MDBSize</td>" + "`r`n"
	Write-Debug "MBXCount:  $MBXCount"
	if ($MBXCount -ne $null)
		{$Write += "<td>$MBXCount</td>" + "`r`n"}
	else
		{$Write += "<td>-</td>" + "`r`n"}
	#Write-Debug $CheckLastBackup
	#Write-Debug $CheckLastBackup.days
	Write-Debug "BackupDays $BackupDays"
	if ($BackupDays -eq "-")
		{$Write += "<td>-</td>" + "`r`n"}
	elseif ($BackupDays -le -2)
		{$Write += "<td bgcolor=red>$LastBackup</td>" + "`r`n"}
	elseif ($BackupDays -eq -1)
		{$Write += "<td bgcolor=yellow>$LastBackup</td>" + "`r`n"}
	elseif ($BackupDays -eq 0)
		{$Write += "<td bgcolor=#BBFBA8>$LastBackup</td>" + "`r`n"}
	else
		{$Write += "<td bgcolor=#BBFBA8>$LastBackup</td>" + "`r`n"}
	
	if ($IncBackupDays -eq "-")
		{$Write += "<td>-</td>" + "`r`n"}
	elseif ($IncBackupDays -le -2)
		{$Write += "<td bgcolor=red>$IncrementalBackup</td>" + "`r`n"}
	elseif ($IncBackupDays -eq -1)
		{$Write += "<td bgcolor=yellow>$IncrementalBackup</td>" + "`r`n"}
	elseif ($IncBackupDays -eq 0)
		{$Write += "<td bgcolor=#BBFBA8>$IncrementalBackup</td>" + "`r`n"}
	else
		{$Write += "<td bgcolor=#BBFBA8>$IncrementalBackup</td>" + "`r`n"}	

	# Write-Debug '$WhiteSpace2: ' $WhiteSpace
	# Write-Debug $WhiteSpace.GetType()
	
	if ($WhiteSpace -is [int] -and $WhiteSpace -ge $Thres30Pct)
		{$Write += "<td bgcolor=red>$WhiteSpace</td>" + "`r`n"}
	elseif ($WhiteSpace -is [int] -and $WhiteSpace -ge $Thres20Pct)
		{$Write += "<td bgcolor=yellow>$WhiteSpace</td>" + "`r`n"}
	elseif ($WhiteSpace -eq "No defrag last day")
		{$Write += "<td bgcolor=yellow>$WhiteSpace</td>" + "`r`n"}
	else
		{$Write += "<td>$WhiteSpace</td>" + "`r`n"}
	$Write += "<td>$DiskCapacity</td>" + "`r`n"
	$Write += "<td>$FreeOnStorage</td>" + "`r`n"
	$FreePct = [int]$FreeOnStoragePct.Remove($FreeOnStoragePct.Length-1,1)
	If ($FreePct -le 5)
		{$Write += "<td bgcolor=red>$FreeOnStoragePct</td>" + "`r`n"}
	ElseIf ($FreePct -lt 15)
		{$Write += "<td bgcolor=yellow>$FreeOnStoragePct</td>" + "`r`n"}
	else
		{$Write += "<td bgcolor=#BBFBA8>$FreeOnStoragePct</td>" + "`r`n"}
	$Write += "</tr>" + "`r`n"
	Return $Write
} #end of Function body

## FOOTER HTML Code
Function Footer() {
$finish = Get-Date
$Write += "</table>" + "`r`n"
$Write += "<br>" + "`r`n"
$Write += "<table class=""style9"">" + "`r`n"
$Write += "<tr>" + "`r`n"
$Write += "<td style=""color #CCCCCC;"" class=""style13"">Finished processing " + $TotalServers + " servers @ " + $finish + " " + $TZ + "</td>" + "`r`n"
$Write += "</tr>" + "`r`n"
$Write += "</table>" + "`r`n"

$Write += "</body></html>" + "`r`n"

$Write += '<p style= "font-family:trebuchet ms;color:#444444;font-size:110%;font-weight:Bold">COMPANY Messaging Report - Exchange Report<p>'
$Write += '<p style= "font-family:trebuchet ms;color:#0085c3;font-size:110%;font-weight:Bold">COMPANY<p>'


Return $Write
} #end of Function Footer


## Main
$Code += Head
#Write-Debug '$Code1: ' $Code
#$Code += Header
#Write-Debug '$Code2: ' $Code

## 
# Build list of Exchange Servers
## Function Main

$skiplist = @()
Import-Csv $ServerSkipList | ForEach-Object {$skiplist += $_.Servername}


#$exchange2007servers = @(Get-MailboxServer | sort) #Get-ExchangeServer |where-object {$_.admindisplayversion.major -eq 8 -and $_.IsMailboxServer -eq $true }
#$exchange2003servers = @(Get-ExchangeServer | where-object {$_.admindisplayversion.major -ne 8 } | sort)

#skiplist test added 
$exchange2007servers = @(Get-MailboxServer | Where-Object {$skiplist -notcontains $_.name} | Sort-Object Name | Select-Object Name) 
$exchange2003servers = @(Get-ExchangeServer | Where-Object {$skiplist -notcontains $_.name -and $_.admindisplayversion.major -ne 8 } | Sort-Object Name | Select-Object Name)
$exchangeServers = $exchange2003servers + $exchange2007servers | Sort Name
#$exchangeServers = "" | select Name
#$exchangeServers.name = "COMPANY-msg-821"
## Function Main Exchange 2007
#if ($false) {  #skip
foreach ($server in $exchangeServers) {
	Write-Debug "======================"
	Write-Debug $server.name
	Write-Debug "Getting Mailbox Databases"
	$ServerDatabases = Get-MailboxDatabase -Status -server $server.name | sort
		If ($ServerDatabases -ne $null) {
			$Code += Header
			$Volumes = Get-WMIObject Win32_Volume -computerName $server.name 
			
			
			
			Write-Debug "Getting 1221 events"
			
			
			$WhiteSpaceEvents = Import-Csv "$($whiteSpaceReportPath)\$($whiteSpaceReportname)"
			
			foreach ($MailboxDatabase in $ServerDatabases) {
				Write-Debug "*****************"
				$Thres20Pct = $null
				$Thres30Pct = $null
				$DiskCapacity = $null
				$FreeOnStorage = $null
				$FreeOnStoragePercent = $null
				$MDBSize = $null
				$WhiteSpace =$null
				$MBXCount = $null
				$BackupDays = $null
				
				Write-Debug "DatabaseName:  $MailboxDatabase.name"
				###diff here
				$edbfilepath = $MailboxDatabase.edbfilepath.pathname
				#$edbpath = $edbfilepath.ToUpper()
				Write-Debug "edbpath:  $($edbfilepath.ToUpper())"
				#convert absolute path to UNC path
				$path = "\\" + $server.name + "\" + $MailboxDatabase.EdbFilePath.PathName.Replace(":","$")
				Write-Debug "path:   $path"
				$dbsize = Get-ChildItem $path
				Write-Debug "dbsize:   $dbsize"
				####mailboxcount different
				#$mailboxcount = Get-MailboxStatistics -database "$MailboxDatabase" | measure-object
				$mailboxcount = Get-MailboxDatabase -Server $server.name | where { $_.Name -eq $MailboxDatabase.name } | Get-Mailbox -ResultSize:Unlimited | measure-object
				
				#$start = $path.LastIndexOf('\')
				#$dbName = $path.Substring($start +1).remove($path.Substring($start +1).length -4)
				#$mailboxpath = "$server\$MailboxDatabase.name"
				#Write-Debug "DatabaseFile: " $mailboxpath
				$edbpathlength = 0
				#$volumes = Get-WMIObject Win32_Volume -computerName $server  #does not need to run more than once per server
				foreach ($volItem in $Volumes) {
					If ( $edbfilepath.ToUpper().startswith($volItem.name.ToUpper()) ) {
						Write-Debug ($volItem.name.ToUpper() +"   -   "+ $edbfilepath.startswith($volItem.name) +"   -   "+ $volItem.name.length)			
						If ($edbpathlength -lt $volItem.name.length) {	
							$edbpathlength = $volItem.name.length
							$edbfilepathvolumename = $volItem.name
							$edbfilepathfreespace =  $volItem.freespace
							$edbfilepathcapacity = $volItem.capacity
						}
					}
				}
				
				$VolumeName = $edbfilepathvolumename
				$FreeOnStorage = "{0:N2}" -f ($edbfilepathfreespace/1GB)
				$DiskCapacity = "{0:N2}" -f ($edbfilepathcapacity/1GB)

				Write-Debug "VolumeName:  $VolumeName"
				Write-Debug "DiskCapacity:  $DiskCapacity"
				Write-Debug "FreeOnStorage:  $FreeOnStorage"
				$FreeOnStoragePercent = "{0:P2}" -f ([int] $FreeOnStorage / [int] $DiskCapacity)
				Write-Debug "FreeOnStoragePercent:  $FreeOnStoragePercent"
				$MDBSize = ("{0:N2}" -f ($dbsize.Length/1GB))
				Write-Debug "MDBSize:  $MDBSize"
				$MBXCount = $mailboxcount.count
				$LastBackup = $MailboxDatabase.LastFullBackup
				Write-Debug "LastBackup:  $LastBackup"
				$IncrementalBackup = $MailboxDatabase.LastIncrementalBackup
				Write-Host '$IncrementalBackup: ' $IncrementalBackup
				$WhiteSpace = Get-WhiteSpace $server $MailboxDatabase.name $WhiteSpaceEvents
				If ($LastBackup -ne $null) { 
					$CheckLastBackup = New-TimeSpan $(get-date) $($LastBackup)
					$BackupDays = $CheckLastBackup.days
				}	
				else { $BackupDays = "-"}
				If ($IncrementalBackup -ne $null) { 
					$CheckLastBackup = New-TimeSpan $(get-date) $($IncrementalBackup)
					$IncBackupDays = $CheckLastBackup.days
				}
				else { $IncBackupDays = "-"}
				$Thres20Pct = [int] $MDBSize * 1024 * .2
				$Thres30Pct = [int] $MDBSize * 1024 * .3
				Write-Debug "Thres20Pct:  $Thres20Pct"
				Write-Debug "Thres30Pct:  $Thres30Pct"
				
				$Code += (body $MailboxDatabase $MDBSize $MBXCount $BackupDays $IncBackupDays $WhiteSpace $DiskCapacity $FreeOnStorage $FreeOnStoragePercent)
			}
		}
	#Write-Debug "Press any key to continue ..."
	#$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	#if ($x.character -eq "x") {exit}
}  # End of Main Exchange 2007/2003 loop

$TotalServers = $exchangeServers.Count

## Wrap Up
$Code += Footer
# Write-Debug '$Code3: ' $Code
#Write-Debug $Code

$Code | out-file $filelocation

$body = $Code


Write-Debug "Done"

If ($SendEmail)   { #switch parameter passed as parameter
	Send-Email $OneOffEmailAddr  $DefaultTo  $DefaultCc  $DefaultBcc  $DefaultFrom  $Subject  $body  $attach  $SMTPHost
	Write-Debug "Email Sent!"
}
else {
	Write-Debug "No Email!"
}

set-psdebug -Off