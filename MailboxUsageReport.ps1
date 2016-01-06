############################################
# This script uses EWS to report on mailbox usage
# Author: 
# Language: Powershell w/Exchange 2007 - tested with exchange 2007 sp3 ru5
# Date: 05/07/2012
############################################


function GetUnreadInboxMails([string] $mailboxname) 
{
###
Write-host "User : " $mailboxname
#Add-Content -Value "User: $mailboxname" -Path $logfile
#$MailboxName = $args[0]
$MailDate = [system.DateTime]::Now.AddDays(-90)

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()

$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
$InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$Sfir = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $false)
$Sflt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $MailDate)
$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfCollection.add($Sfir)
$sfCollection.add($Sflt)
$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(500)
$frFolderResult = $InboxFolder.FindItems($sfCollection,$view)
"Number of Unread Email for the Last 90 Days : " + $frFolderResult.Items.Count
[string]$numberofunread = $frFolderResult.Items.Count
#Add-Content -Value "Number of Unread Email for the Last 90 Days :  $numberofunread" -Path $logfile
if ($frFolderResult.Items.Count -ne 0){
	#"Last Unread Subject : " + $frFolderResult.Items[0].Subject
	"Last Unread DateTime : " + $frFolderResult.Items[499].DateTimeReceived
	[string]$lastunreaddate = $frFolderResult.Items[499].DateTimeReceived
	#Add-Content -Value "Last Unread DateTime :  $lastunreaddate" -Path $logfile
}
$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(1)
$frFolderResult = $InboxFolder.FindItems($view)
if ($frFolderResult.Items.Count -ne 0){
	#"Last Recieved Subject : " + $frFolderResult.Items[0].Subject
	"Last Recieved DateTime : " + $frFolderResult.Items[0].DateTimeReceived
	[string]$lastreceieveddate = $frFolderResult.Items[0].DateTimeReceived
	#Add-Content -Value "Last Recieved DateTime :  $lastreceieveddate" -Path $logfile
}
$sfview = new-object Microsoft.Exchange.WebServices.Data.ItemView(1)
$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$MailboxName)
$SentItemsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$srFolderResult = $SentItemsFolder.FindItems($sfview)
if ($srFolderResult.Items.Count -ne 0){
	#"Last Sent Subject : " + $srFolderResult.Items[0].Subject
	"Last Sent DateTime : " + $srFolderResult.Items[0].DateTimeReceived
	[string]$lastsentdatetime = $srFolderResult.Items[0].DateTimeReceived
	#Add-Content -Value "Last Sent DateTime :  $lastsentdatetime" -Path $logfile

}
$cfview = new-object Microsoft.Exchange.WebServices.Data.ItemView(1)
$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)
$ContactsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$cfFolderResult = $ContactsFolder.FindItems($sfview)
if ($srFolderResult.Items.Count -ne 0){
	#"Last Contact Created: " + $cfFolderResult.Items[0].Subject
	"Last Contact CreatedTime : " + $cfFolderResult.Items[0].DateTimeReceived
	[string]$lastcontactctime = $cfFolderResult.Items[0].DateTimeReceived
	#Add-Content -Value "Last Contact CreatedTime :  $lastcontactctime" -path $logfile
}
$apview = new-object Microsoft.Exchange.WebServices.Data.ItemView(1)
$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)
$CalendarFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$cfFolderResult = $CalendarFolder.FindItems($apview)
if ($srFolderResult.Items.Count -ne 0){
	#"Last Appointment Created: " + $cfFolderResult.Items[0].Subject
	"Last Appointment CreatedTime : " + $cfFolderResult.Items[0].DateTimeReceived
	[string]$LastapptCtime =  $cfFolderResult.Items[0].DateTimeReceived
	#Add-Content -Value "Last Appointment CreatedTime :  $LastapptCtime " -Path $logfile
}
###
Add-Content -Value "$mailboxname`t$numberofunread`t$lastunreaddate`t$lastsentdatetime`t$lastcontactctime`t$LastapptCtime" -Path $logfile
}

$inputfile = "c:\scripts\inuse\allemailaddresses.csv"
$logfile = "c:\scripts\inuse\unreadmailboxes07212012.xls"

Add-Content -Value "User`tNumber of Unread Email for the Last 90 Days`tLast Unread DateTime`tLast Sent DateTime`tLast Contact CreatedTime`tLast Appointment CreatedTime" -path $logfile

$addresses = Import-Csv $inputFile 
$addresses | % {GetUnreadInboxMails($_.Mailboxname)}  
