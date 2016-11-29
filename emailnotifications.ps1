<#
Title:               Email Reporting Script
Author:              James Auchterlonie.
Creation Date:       25/11/16
Last Edit:           28/11/16
Last Editor:         James Auchterlonie.
Script Function:     This script generates a html email containing information on an array of emails.

Version History:
-----------------------------------------------------
1 | Complete . Lines Edited: 1->111
        -Select previous month's emails
        -Select information from exchange records.
        -Send Html Email
#>

#variable declaration
$emails = @("example@domain.com","example@domain.com","example@domain.com","example@domain.com","example@domain.com" , "example@domain.com", "example@domain.com" , "example@domain.com")
$sent = $received = $size ="0"
$user = "Chris"
$emailcontent =""

#declare smtp email values
$mailTo = "receiver@domain.com"
$mailFrom = "sender@domain.com"
$todaydate = Get-Date -DisplayHint date
$subject = "Email Reporting: $todaydate"
$smtpServer = "smtp address"


# Load Exchange Snap-in
if ( (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction silentlycontinue) -eq $null )
{
	Write-Host -ForegroundColor Yellow "Loading modules...Please wait..."
	add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010
}
# Load connect functions
$global:exbin = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath + "bin\"
. $global:exbin"CommonConnectFunctions.ps1"
. $global:exbin"ConnectFunctions.ps1"

# Connect to an exchange server
Connect-ExchangeServer -Auto -AllowClobber
# Import AD module
if (-not (Get-Module ActiveDirectory | Measure-Object).count)
{
	Import-Module ActiveDirectory
}

#work out previous month dates
$date = get-date
$numdays = $date.Day
$prevmonthlastday = $date.AddDays(-$numdays)
$numdays = $prevmonthlastday.Day
$prevmonthfirstday = $prevmonthlastday.AddDays( -$numdays + 1)


#declare email style using css
$emailcontent += '<style>'
$emailcontent += 'body{'
$emailcontent += '  font-family: Arial, Helvetica, sans-serif;'
$emailcontent += '  font-size: 12;'
$emailcontent += '}'
$emailcontent += 'table, th, td {'
$emailcontent += '    border: 1px solid black;'
$emailcontent += '    border-collapse: collapse;'
$emailcontent += '    padding: 5px;'
$emailcontent += '    font-family: Arial, Helvetica, sans-serif;'
$emailcontent += '    font-size: 11;'
$emailcontent += '}'
$emailcontent += '</style>'

#declare email content using html
$emailcontent += '<body>'
$emailcontent +="<p>Morning $user,</p>"
$emailcontent +='<p>Please find below information regarding email usage over the past month.</p>'
$emailcontent += '	<table>'
$emailcontent += '		<tr style="background-color: #3399ff; color:white; font-weight: bold;">'
$emailcontent += '		  <td> Email Address </td>'
$emailcontent += '		  <td> Sent </td>'
$emailcontent += '		  <td> Received. </td>'
$emailcontent += '		  <td> Size (MB) </td>'
$emailcontent += '		</tr>'

#generate table rows for each entry in the $emails array.
foreach($email in $emails){
    $intSent = $intRec=0
    $size = Get-MailboxStatistics $email | select TotalItemSize
    Get-MessageTrackingLog -ResultSize Unlimited -Start “$prevmonthfirstday” -End “$prevmonthlastday” -Sender "$email" -EventID RECEIVE | ForEach { $intSent++ }
    Get-MessageTrackingLog -ResultSize Unlimited -Start "$prevmonthfirstday" -End "$prevmonthlastday" -Recipients "$email" -EventID DELIVER | ForEach { $intRec++ }
    $emailcontent +="    <tr>"
	$emailcontent +="	      <td>$email</td>"
	$emailcontent +="	      <td>$intSent</td>"
	$emailcontent +="	      <td>$intRec</td>"
	$emailcontent +="	      <td>$size</td>"
	$emailcontent +="	 </tr>"
}

#declare final email content
$emailcontent +='</table>'
$emailcontent +='  <p>Regards,</p>'
$emailcontent +='  <p>'
$emailcontent +='  Sender Name   '
$emailcontent +='  <p>'
$emailcontent +='</body>'


#send email to user.
Send-MailMessage -To "$mailTo" -From "$mailFrom" -SmtpServer "$smtpServer" -Subject "$subject" -BodyAsHtml $emailContent
