#requires -version 7.0
<#
.SYNOPSIS
  Create iCal / vCalendar file for use across all devices; used internally for Holiday/PayDay/TimeSheet reminders.
.DESCRIPTION
  This script imports a CSV and generates an ICS file for distribution. Leverages both:
  RFC2445:
  https://www.ietf.org/rfc/rfc2445.txt
  and Microsofts iCal to Appointment conversion tags:
  https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcical/e7a90cf6-fd81-4dfd-8dee-bcbde9c4fe05
.PARAMETER calName
  REQUIRED. Specifies the name of the Calendar as it appears in the Outlook list.
.PARAMETER categories
  Optional. Specifies the Outlook Category the entries will be designated. 
.PARAMETER startDate
  Optional. The date of the event, in format yyyyMMdd.
.PARAMETER endDate
  Optional. The end date of the event (usually startDate++), in format yyyyMMdd.
.PARAMETER eventSubject
  Optional. The name of the event, i.e. Christmas Day.
.PARAMETER eventDesc
  Optional. Further description of the event.
.PARAMETER eventLocation
  Optional. Defaults to United States.
.PARAMETER csvFileInPath
  REQUIRED. Path to CSV inout file.
.PARAMETER icsFileOutPath
  REQUIRED. Path to ICS (iCal) output file.
.INPUTS
  CSV file or Custom PSObjects from pipeline by PropertyName
.OUTPUTS
  Properly formatted ICS file for distribution  
.NOTES
  Version:        1.2
  Author:         Dave Negrotto
  Creation Date:  12.20.2019
  Last Update:    07.02.2021
  Purpose/Change: Add X-WR-RELCALID attribute for Outlook support
  Extension of Justin Braun's work: https://justinbraun.com/2018/01/powershell-dynamic-generation-of-an-ical-vcalendar-ics-format-file/
  Example CSVs included in Repo
.EXAMPLE
  Default execution creates a CompanyName Holiday calendar.
  .\New-IcsFile.ps1 -calName 'CompanyName Holidays' -csvFileInPath C:\CompanyName_Year_Holiday_Calendar.csv -icsFileOutPath C:\CompanyName_Year_Holiday_Calendar.ics
.EXAMPLE
  Creating a Payday entry.
  .\New-IcsFile.ps1 -calName 'CompanyName Paydays' -csvFileInPath C:\CompanyName_Year_Payday_Calendar.csv -icsFileOutPath C:\CompanyName_Year_Payday_Calendar.ics
.EXAMPLE
  Creating a Timesheets entry.
  .\New-IcsFile.ps1 -calName 'CompanyName Timesheets' -csvFileInPath C:\CompanyName_Year_Timesheet_Calendar.csv -icsFileOutPath C:\CompanyName_Year_Timesheets_Calendar.ics 
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName)][string][ValidateSet('CompanyName Holidays', 'CompanyName Paydays', 'CompanyName Timesheets')]$calName = 'CompanyName Holidays',
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName)][string][ValidateSet('Holiday', 'Payday', 'Timesheets')]$categories = 'Holiday',
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName, HelpMessage="yyyyMMdd")][string]$startDate,
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName, HelpMessage="yyyyMMdd")][string]$endDate,
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName)][string]$eventSubject,
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName)][string]$eventDesc,
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName)][string]$eventLocation = 'United States',
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName, HelpMessage="c:\CompanyName_Calendar.csv")][string]$csvFileInPath,
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName, HelpMessage="c:\CompanyName_Calendar.ics")][string]$icsFileOutPath
)

# Custom date formats that we want to use
$longDateFormat = "yyyyMMddTHHmmssZ"


# Instantiate .NET StringBuilder
$sb = [System.Text.StringBuilder]::new()
  
# Fill in ICS/iCalendar properties based on RFC2445
[void]$sb.AppendLine('BEGIN:VCALENDAR')
[void]$sb.AppendLine('VERSION:2.0')
[void]$sb.AppendLine('METHOD:PUBLISH')
[void]$sb.AppendLine('PRODID:-//CompanyName//PowerShell ICS Creator//EN')
[void]$sb.AppendLine('X-WR-CALNAME:' + $calName)

if (Test-Path -Path $csvFileInPath) { # If input file exists, START LOOP FOR EACH EVENT
	Import-Csv $csvFileInPath | ForEach-Object {
	  $categories = $_.categories
	  $startDate = $_.startDate
	  $endDate = $_.endDate
	  $eventSubject = $_.eventSubject
	  $eventDesc =  $_.eventDesc
	  $eventLocation = $_.eventLocation

	  [void]$sb.AppendLine('BEGIN:VEVENT')
	  [void]$sb.AppendLine('CLASS:PUBLIC')
	  [void]$sb.AppendLine("CATEGORIES:" + $categories)
	  [void]$sb.AppendLine("UID:" + [guid]::NewGuid())
	  [void]$sb.AppendLine("CREATED:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
	  [void]$sb.AppendLine("DTSTAMP:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
	  [void]$sb.AppendLine("LAST-MODIFIED:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
	  [void]$sb.AppendLine("SEQUENCE:0")
	  [void]$sb.AppendLine("PRIORITY:5")
	  [void]$sb.AppendLine("DTSTART;VALUE=DATE:" + $startDate)
	  [void]$sb.AppendLine("DTEND;VALUE=DATE:" + $endDate)
	  [void]$sb.AppendLine("DESCRIPTION:" + $eventDesc)
	  [void]$sb.AppendLine("SUMMARY;LANGUAGE=en-us:" + $eventSubject)
	  [void]$sb.AppendLine("LOCATION:" + $eventLocation)  
	  [void]$sb.AppendLine("TRANSP:TRANSPARENT")  
	  [void]$sb.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:FREE")
	  [void]$sb.AppendLine("X-MICROSOFT-CDO-IMPORTANCE:1")
	  [void]$sb.AppendLine("X-MICROSOFT-DISALLOW-COUNTER:FALSE")
	  [void]$sb.AppendLine("X-MS-OLK-AUTOFILLLOCATION:FALSE")
	  [void]$sb.AppendLine("X-MS-OLK-AUTOSTARTCHECK:FALSE")
	  [void]$sb.AppendLine("X-MS-OLK-CONFTYPE:0")
	  [void]$sb.AppendLine('END:VEVENT')
	} # END LOOP
	
}
else {
	[void]$sb.AppendLine('BEGIN:VEVENT')
	[void]$sb.AppendLine('CLASS:PUBLIC')
	[void]$sb.AppendLine("CATEGORIES:" + $categories)
	[void]$sb.AppendLine("UID:" + [guid]::NewGuid())
	[void]$sb.AppendLine("CREATED:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
	[void]$sb.AppendLine("DTSTAMP:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
	[void]$sb.AppendLine("LAST-MODIFIED:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
	[void]$sb.AppendLine("SEQUENCE:0")
	[void]$sb.AppendLine("PRIORITY:5")
	[void]$sb.AppendLine("DTSTART;VALUE=DATE:" + $startDate)
	[void]$sb.AppendLine("DTEND;VALUE=DATE:" + $endDate)
	[void]$sb.AppendLine("DESCRIPTION:" + $eventDesc)
	[void]$sb.AppendLine("SUMMARY;LANGUAGE=en-us:" + $eventSubject)
	[void]$sb.AppendLine("LOCATION:" + $eventLocation)  
	[void]$sb.AppendLine("TRANSP:TRANSPARENT")  
	[void]$sb.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:FREE")
	[void]$sb.AppendLine("X-MICROSOFT-CDO-IMPORTANCE:1")
	[void]$sb.AppendLine("X-MICROSOFT-DISALLOW-COUNTER:FALSE")
	[void]$sb.AppendLine("X-MS-OLK-AUTOFILLLOCATION:FALSE")
	[void]$sb.AppendLine("X-MS-OLK-AUTOSTARTCHECK:FALSE")
	[void]$sb.AppendLine("X-MS-OLK-CONFTYPE:0")
	[void]$sb.AppendLine('END:VEVENT')
}

[void]$sb.AppendLine('END:VCALENDAR')

# Output ICS File
$sb.ToString() | Out-File $icsFileOutPath # Ensure outFile is UTF-8 no BOM
