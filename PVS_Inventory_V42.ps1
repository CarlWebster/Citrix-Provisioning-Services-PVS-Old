#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
	Creates a complete inventory of a Citrix PVS 5.x, 6.x or 7.x farm using Microsoft Word 2010 or 2013.
.DESCRIPTION
	Creates a complete inventory of a Citrix PVS 5.x, 6.x or 7.x farm using Microsoft Word and PowerShell.
	Creates a Word document named after the PVS 5.x, 6.x or 7.x farm.
	Document includes a Cover Page, Table of Contents and Footer.
	Version 4 and later include support for the following language versions of Microsoft Word:
		Catalan
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
	If either registry key does not exist and this parameter is not specified, the report will
	not contain a Company Name on the cover page.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010 and 2013 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013. Doesn't work in 2013, mostly works in 2010 but 
						Subtitle/Subject & Author fields need to be moved 
						after title box is moved up)
		Banded (Word 2013. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013. Works)
		Filigree (Word 2013. Works)
		Grid (Word 2010/2013.Works in 2010)
		Integral (Word 2013. Works)
		Ion (Dark) (Word 2013. Top date doesn't fit, box needs to be manually resized or font 
						changed to 8 point)
		Ion (Light) (Word 2013. Top date doesn't fit, box needs to be manually resized or font 
						changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit, box needs to be manually resized or font 
					changed to 14 point)
		Retrospect (Word 2013. Works)
		Semaphore (Word 2013. Works)
		Sideline (Word 2010/2013. Doesn't work in 2013, works in 2010)
		Slice (Dark) (Word 2013. Doesn't work)
		Slice (Light) (Word 2013. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013. Works)
		Whisp (Word 2013. Works)
		
	Default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
	This parameter is reserved for a future update and no useable output is created at this time.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
	This parameter is reserved for a future update and no output is created at this time.
.PARAMETER Hardware
	Use WMI to gather hardware information on: Computer System, Disks, Processor and Network Interface Cards
	This parameter is disabled by default.
.PARAMETER AdminAddress
	Specifies the name of a PVS server that the PowerShell script will connect to. 
.PARAMETER User
	Specifies the user used for the AdminAddress connection. 
.PARAMETER Domain
	Specifies the domain used for the AdminAddress connection. 
.PARAMETER Password
	Specifies the password used for the AdminAddress connection. 
.PARAMETER StartDate
	Start date, in MM/DD/YYYY format, for the Audit Trail report.
	Default is today's date minus seven days.
.PARAMETER EndDate
	End date, in MM/DD/YYYY format, for the Audit Trail report.
	Default is today's date.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be ReportName_2014-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.EXAMPLE
	PS C:\PSScript > .\PVS_Inventory_V42.ps1
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator
	AdminAddress = LocalHost

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	LocalHost for AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\PVS_Inventory_V42.ps1 -PDF 
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator
	AdminAddress = LocalHost

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	LocalHost for AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\PVS_Inventory_V42.ps1 -TEXT

	This parameter is reserved for a future update and no output is created at this time.
	
	Will use all default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	LocalHost for AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\PVS_Inventory_V42.ps1 -HTML

	This parameter is reserved for a future update and no output is created at this time.
	
	Will use all default values and save the document as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	LocalHost for AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\PVS_Inventory_V42.ps1 -Hardware 
	
	Will use all Default values and add additional information for each server about its hardware.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\PVS_Inventory_V42.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\PVS_Inventory_V42.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster" -AdminAddress PVS1

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
		PVS1 for AdminAddress.
.EXAMPLE
	PS C:\PSScript .\PVS_Inventory_V42.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster" -AdminAddress PVS1 -User cwebster -Domain WebstersLab -Password Abc123!@#

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
		PVS1 for AdminAddress.
		cwebster for User.
		WebstersLab for Domain.
		Abc123!@# for Password.
.EXAMPLE
	PS C:\PSScript .\PVS_Inventory_V42.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster" -AdminAddress PVS1 -User cwebster

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
		PVS1 for AdminAddress.
		cwebster for User.
		Script will prompt for the Domain and Password
.EXAMPLE
	PS C:\PSScript > .\PVS_Inventory_V42.ps1 -StartDate "01/01/2014" -EndDate "01/31/2014" 
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator
	AdminAddress = LocalHost

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	LocalHost for AdminAddress.
	Will return all Audit Trail entries from "01/01/2014" through "01/31/2014".
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word or PDF document.
.NOTES
	NAME: PVS_Inventory_V42.ps1
	VERSION: 4.21
	AUTHOR: Carl Webster (with a lot of help from Michael B. Smith, Jeff Wouters and Iain Brighton)
	LASTEDIT: July 8, 2015
#>


#thanks to @jeffwouters for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="Text",Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(ParameterSetName="HTML",Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Hardware=$False, 

	[parameter(Mandatory=$False)] 
	[Datetime]$StartDate = ((Get-Date -displayhint date).AddDays(-7)),

	[parameter(Mandatory=$False)] 
	[Datetime]$EndDate = (Get-Date -displayhint date),
	
	[parameter(Mandatory=$False)] 
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$AdminAddress="",

	[parameter(Mandatory=$False)] 
	[string]$Domain="",

	[parameter(Mandatory=$False)] 
	[string]$User="",

	[parameter(Mandatory=$False)] 
	[string]$Password="",
	
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username

	)


#Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#This script written for "Benji", March 19, 2012
#Thanks to Michael B. Smith, Joe Shonk and Stephane Thirion
#for testing and fine-tuning tips 

#Version 4.2
#	Cleanup the script's parameters section
#	Code cleanup and standardization with the master template script
#	Requires PowerShell V3 or later
#	Removed support for Word 2007
#	Word 2007 references in help text removed
#	Cover page parameter now states only Word 2010 and 2013 are supported
#	Most Word 2007 references in script removed:
#		Function ValidateCoverPage
#		Function SetupWord
#		Function SaveandCloseDocumentandShutdownWord
#	Function CheckWord2007SaveAsPDFInstalled removed
#	If Word 2007 is detected, an error message is now given and the script is aborted
#	Cleanup Word table code for the first row and background color
#	Add Iain Brighton's Word table functions
#	Move Appendix A and B tables to new table function
#	Move hardware info to new table functions
#	Move audit trail info to new table functions
#	Add parameters for MSWord, Text and HTML for future updates
#
#Version 4.21
#	Add writeCacheType 9 (Cache to Device RAM with overflow to hard disk) for PVS 7.x
#	Remove writeCacheType 3 and 5 from PVS 6 and 7
#	Updated help text
#	Updated hardware inventory code

Set-StrictMode -Version 2

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

If($PDF -eq $Null)
{
	$PDF = $False
}
If($Text -eq $Null)
{
	$Text = $False
}
If($MSWord -eq $Null)
{
	$MSWord = $False
}
If($HTML -eq $Null)
{
	$HTML = $False
}
If($AddDateTime -eq $Null)
{
	$AddDateTime = $False
}
If($Hardware -eq $Null)
{
	$Hardware = $False
}
If($ComputerName -eq $Null)
{
	$ComputerName = "LocalHost"
}

If(!(Test-Path Variable:PDF))
{
	$PDF = $False
}
If(!(Test-Path Variable:Text))
{
	$Text = $False
}
If(!(Test-Path Variable:MSWord))
{
	$MSWord = $False
}
If(!(Test-Path Variable:HTML))
{
	$HTML = $False
}
If(!(Test-Path Variable:AddDateTime))
{
	$AddDateTime = $False
}
If(!(Test-Path Variable:Hardware))
{
	$Hardware = $False
}
If(!(Test-Path Variable:ComputerName))
{
	$ComputerName = "LocalHost"
}

If($MSWord -eq $Null)
{
	If($Text -or $HTML -or $PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
ElseIf($Text)
{
	Write-Verbose "$(Get-Date): Text is set"
}
ElseIf($HTML)
{
	Write-Verbose "$(Get-Date): HTML is set"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($MSWord -eq $Null)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($PDF -eq $Null)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	ElseIf($Text -eq $Null)
	{
		Write-Verbose "$(Get-Date): Text is Null"
	}
	ElseIf($HTML -eq $Null)
	{
		Write-Verbose "$(Get-Date): HTML is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
		Write-Verbose "$(Get-Date): PDF is $($PDF)"
		Write-Verbose "$(Get-Date): Text is $($Text)"
		Write-Verbose "$(Get-Date): HTML is $($HTML)"
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $($CoName)"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[long]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[int]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdFormatDocumentDefault = 16
	[int]$wdSaveFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 

	[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
}

#region code for -hardware switch
Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	

	#Get Computer info
	Write-Verbose "$(Get-Date): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date): `t`t`tHardware information"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteWordLine 4 0 "General Computer"
	}
	ElseIf($Text)
	{
		Line 0 "Computer Information: $($RemoteComputerName)"
		Line 1 "General Computer"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Computer Information: $($RemoteComputerName)"
	}
	
	[bool]$GotComputerItems = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Results -ne $Null)
	{
		$ComputerItems = $Results | Select Manufacturer, Model, Domain, `
		@{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}, `
		NumberOfProcessors, NumberOfLogicalProcessors
		$Results = $Null

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
			Line 2 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Computer information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Computer information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Drive(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Drive(s)"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Drive(s)"
	}

	[bool]$GotDrives = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Results -ne $Null)
	{
		$drives = $Results | Select caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				OutputDriveItem $drive
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Drive information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Drive information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
	}
	

	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Processor(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Processor(s)"
	}
	ElseIf($HTML)
	{
	}

	[bool]$GotProcessors = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Results -ne $Null)
	{
		$Processors = $Results | Select availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			OutputProcessorItem $processor
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Processor information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Processor information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Interface(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Network Interface(s)"
	}
	ElseIf($HTML)
	{
	}

	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch
	{
		$Results
	}

	If($? -and $Results -ne $Null)
	{
		$Nics = $Results | Where {$_.ipaddress -ne $Null}
		$Results = $Null

		If($Nics -eq $Null ) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where {$_.index -eq $nic.index}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $ThisNic -ne $Null)
				{
					OutputNicItem $Nic $ThisNic
				}
				ElseIf(!$?)
				{
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "Error retrieving NIC information"
						Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
						Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
						Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
						Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date): No results Returned for NIC information"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "No results Returned for NIC information"
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Error retrieving NIC configuration information"
			Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for NIC configuration information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 0 0 ""
	}

	$Results = $Null
	$ComputerItems = $Null
	$Drives = $Null
	$Processors = $Null
	$Nics = $Null
}

Function OutputComputerItem
{
	Param([object]$Item)
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ItemInformation = @()
		$ItemInformation += @{ Data = "Manufacturer"; Value = $Item.manufacturer; }
		$ItemInformation += @{ Data = "Model"; Value = $Item.model; }
		$ItemInformation += @{ Data = "Domain"; Value = $Item.domain; }
		$ItemInformation += @{ Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }
		$ItemInformation += @{ Data = "Physical Processors (sockets)"; Value = $Item.NumberOfProcessors; }
		$ItemInformation += @{ Data = "Logical Processors (cores w/HT)"; Value = $Item.NumberOfLogicalProcessors; }
		$Table = AddWordTable -Hashtable $ItemInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	ElseIf($Text)
	{
		Line 2 "Manufacturer`t: " $Item.manufacturer
		Line 2 "Model`t`t: " $Item.model
		Line 2 "Domain`t`t: " $Item.domain
		Line 2 "Total Ram`t: $($Item.totalphysicalram) GB"
		Line 2 "Physical Processors (sockets): " $Item.NumberOfProcessors
		Line 2 "Logical Processors (cores w/HT): " $Item.NumberOfLogicalProcessors
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Manufacturer",($htmlsilver -bor $htmlbold),$Item.manufacturer,$htmlwhite)
		$rowdata += @(,('Model',($htmlsilver -bor $htmlbold),$Item.model,$htmlwhite))
		$rowdata += @(,('Domain',($htmlsilver -bor $htmlbold),$Item.domain,$htmlwhite))
		$rowdata += @(,('Total Ram',($htmlsilver -bor $htmlbold),"$($Item.totalphysicalram) GB",$htmlwhite))
		$rowdata += @(,('Physical Processors (sockets)',($htmlsilver -bor $htmlbold),$Item.NumberOfProcessors,$htmlwhite))
		$rowdata += @(,('Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Item.NumberOfLogicalProcessors,$htmlwhite))

		$msg = "General Computer"
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputDriveItem
{
	Param([object]$Drive)
	
	$xDriveType = ""
	Switch ($drive.drivetype)
	{
		0	{$xDriveType = "Unknown"}
		1	{$xDriveType = "No Root Directory"}
		2	{$xDriveType = "Removable Disk"}
		3	{$xDriveType = "Local Disk"}
		4	{$xDriveType = "Network Drive"}
		5	{$xDriveType = "Compact Disc"}
		6	{$xDriveType = "RAM Disk"}
		Default {$xDriveType = "Unknown"}
	}
	
	$xVolumeDirty = ""
	If(![String]::IsNullOrEmpty($drive.volumedirty))
	{
		If($drive.volumedirty)
		{
			$xVolumeDirty = "Yes"
		}
		Else
		{
			$xVolumeDirty = "No"
		}
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $DriveInformation = @()
		$DriveInformation += @{ Data = "Caption"; Value = $Drive.caption; }
		$DriveInformation += @{ Data = "Size"; Value = "$($drive.drivesize) GB"; }
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$DriveInformation += @{ Data = "File System"; Value = $Drive.filesystem; }
		}
		$DriveInformation += @{ Data = "Free Space"; Value = "$($drive.drivefreespace) GB"; }
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$DriveInformation += @{ Data = "Volume Name"; Value = $Drive.volumename; }
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$DriveInformation += @{ Data = "Volume is Dirty"; Value = $xVolumeDirty; }
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation += @{ Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }
		}
		$DriveInformation += @{ Data = "Drive Type"; Value = $xDriveType; }
		$Table = AddWordTable -Hashtable $DriveInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells `
		-Bold `
		-BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	ElseIf($Text)
	{
		Line 2 "Caption`t`t: " $drive.caption
		Line 2 "Size`t`t: $($drive.drivesize) GB"
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			Line 2 "File System`t: " $drive.filesystem
		}
		Line 2 "Free Space`t: $($drive.drivefreespace) GB"
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			Line 2 "Volume Name`t: " $drive.volumename
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			Line 2 "Volume is Dirty`t: " $xVolumeDirty
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			Line 2 "Volume Serial #`t: " $drive.volumeserialnumber
		}
		Line 2 "Drive Type`t: " $xDriveType
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Caption",($htmlsilver -bor $htmlbold),$Drive.caption,$htmlwhite)
		$rowdata += @(,('Size',($htmlsilver -bor $htmlbold),"$($drive.drivesize) GB",$htmlwhite))

		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$rowdata += @(,('File System',($htmlsilver -bor $htmlbold),$Drive.filesystem,$htmlwhite))
		}
		$rowdata += @(,('Free Space',($htmlsilver -bor $htmlbold),"$($drive.drivefreespace) GB",$htmlwhite))
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$rowdata += @(,('Volume Name',($htmlsilver -bor $htmlbold),$Drive.volumename,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$rowdata += @(,('Volume is Dirty',($htmlsilver -bor $htmlbold),$xVolumeDirty,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$rowdata += @(,('Volume Serial Number',($htmlsilver -bor $htmlbold),$Drive.volumeserialnumber,$htmlwhite))
		}
		$rowdata += @(,('Drive Type',($htmlsilver -bor $htmlbold),$xDriveType,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"}
		2	{$xAvailability = "Unknown"}
		3	{$xAvailability = "Running or Full Power"}
		4	{$xAvailability = "Warning"}
		5	{$xAvailability = "In Test"}
		6	{$xAvailability = "Not Applicable"}
		7	{$xAvailability = "Power Off"}
		8	{$xAvailability = "Off Line"}
		9	{$xAvailability = "Off Duty"}
		10	{$xAvailability = "Degraded"}
		11	{$xAvailability = "Not Installed"}
		12	{$xAvailability = "Install Error"}
		13	{$xAvailability = "Power Save - Unknown"}
		14	{$xAvailability = "Power Save - Low Power Mode"}
		15	{$xAvailability = "Power Save - Standby"}
		16	{$xAvailability = "Power Cycle"}
		17	{$xAvailability = "Power Save - Warning"}
		Default	{$xAvailability = "Unknown"}
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $ProcessorInformation = @()
		$ProcessorInformation += @{ Data = "Name"; Value = $Processor.name; }
		$ProcessorInformation += @{ Data = "Description"; Value = $Processor.description; }
		$ProcessorInformation += @{ Data = "Max Clock Speed"; Value = "$($processor.maxclockspeed) MHz"; }
		If($processor.l2cachesize -gt 0)
		{
			$ProcessorInformation += @{ Data = "L2 Cache Size"; Value = "$($processor.l2cachesize) KB"; }
		}
		If($processor.l3cachesize -gt 0)
		{
			$ProcessorInformation += @{ Data = "L3 Cache Size"; Value = "$($processor.l3cachesize) KB"; }
		}
		If($processor.numberofcores -gt 0)
		{
			$ProcessorInformation += @{ Data = "Number of Cores"; Value = $Processor.numberofcores; }
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$ProcessorInformation += @{ Data = "Number of Logical Processors (cores w/HT)"; Value = $Processor.numberoflogicalprocessors; }
		}
		$ProcessorInformation += @{ Data = "Availability"; Value = $xAvailability; }
		$Table = AddWordTable -Hashtable $ProcessorInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t: " $processor.name
		Line 2 "Description`t`t: " $processor.description
		Line 2 "Max Clock Speed`t`t: $($processor.maxclockspeed) MHz"
		If($processor.l2cachesize -gt 0)
		{
			Line 2 "L2 Cache Size`t`t: $($processor.l2cachesize) KB"
		}
		If($processor.l3cachesize -gt 0)
		{
			Line 2 "L3 Cache Size`t`t: $($processor.l3cachesize) KB"
		}
		If($processor.numberofcores -gt 0)
		{
			Line 2 "# of Cores`t`t: " $processor.numberofcores
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			Line 2 "# of Logical Procs (cores w/HT)`t: " $processor.numberoflogicalprocessors
		}
		Line 2 "Availability`t`t: " $xAvailability
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Processor.name,$htmlwhite)
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Processor.description,$htmlwhite))

		$rowdata += @(,('Max Clock Speed',($htmlsilver -bor $htmlbold),"$($processor.maxclockspeed) MHz",$htmlwhite))
		If($processor.l2cachesize -gt 0)
		{
			$rowdata += @(,('L2 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l2cachesize) KB",$htmlwhite))
		}
		If($processor.l3cachesize -gt 0)
		{
			$rowdata += @(,('L3 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l3cachesize) KB",$htmlwhite))
		}
		If($processor.numberofcores -gt 0)
		{
			$rowdata += @(,('Number of Cores',($htmlsilver -bor $htmlbold),$Processor.numberofcores,$htmlwhite))
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$rowdata += @(,('Number of Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Processor.numberoflogicalprocessors,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))

		$msg = "Processor(s)"
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"}
		2	{$xAvailability = "Unknown"}
		3	{$xAvailability = "Running or Full Power"}
		4	{$xAvailability = "Warning"}
		5	{$xAvailability = "In Test"}
		6	{$xAvailability = "Not Applicable"}
		7	{$xAvailability = "Power Off"}
		8	{$xAvailability = "Off Line"}
		9	{$xAvailability = "Off Duty"}
		10	{$xAvailability = "Degraded"}
		11	{$xAvailability = "Not Installed"}
		12	{$xAvailability = "Install Error"}
		13	{$xAvailability = "Power Save - Unknown"}
		14	{$xAvailability = "Power Save - Low Power Mode"}
		15	{$xAvailability = "Power Save - Standby"}
		16	{$xAvailability = "Power Cycle"}
		17	{$xAvailability = "Power Save - Warning"}
		Default	{$xAvailability = "Unknown"}
	}

	$xIPAddress = @()
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress += "$($IPAddress)"
	}

	$xIPSubnet = @()
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet += "$($IPSubnet)"
	}

	If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = @()
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder += "$($DNSDomain)"
		}
	}
	
	If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = @()
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder += "$($DNSServer)"
		}
	}

	$xdnsenabledforwinsresolution = ""
	If($nic.dnsenabledforwinsresolution)
	{
		$xdnsenabledforwinsresolution = "Yes"
	}
	Else
	{
		$xdnsenabledforwinsresolution = "No"
	}
	
	$xTcpipNetbiosOptions = ""
	Switch ($nic.TcpipNetbiosOptions)
	{
		0	{$xTcpipNetbiosOptions = "Use NetBIOS setting from DHCP Server"}
		1	{$xTcpipNetbiosOptions = "Enable NetBIOS"}
		2	{$xTcpipNetbiosOptions = "Disable NetBIOS"}
		Default	{$xTcpipNetbiosOptions = "Unknown"}
	}
	
	$xwinsenablelmhostslookup = ""
	If($nic.winsenablelmhostslookup)
	{
		$xwinsenablelmhostslookup = "Yes"
	}
	Else
	{
		$xwinsenablelmhostslookup = "No"
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $NicInformation = @()
		$NicInformation += @{ Data = "Name"; Value = $ThisNic.Name; }
		If($ThisNic.Name -ne $nic.description)
		{
			$NicInformation += @{ Data = "Description"; Value = $Nic.description; }
		}
		$NicInformation += @{ Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }
		$NicInformation += @{ Data = "Manufacturer"; Value = $Nic.manufacturer; }
		$NicInformation += @{ Data = "Availability"; Value = $xAvailability; }
		$NicInformation += @{ Data = "Physical Address"; Value = $Nic.macaddress; }
		If($xIPAddress.Count -gt 1)
		{
			$NicInformation += @{ Data = "IP Address"; Value = $xIPAddress[0]; }
			$NicInformation += @{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }
			$NicInformation += @{ Data = "Subnet Mask"; Value = $xIPSubnet[0]; }
			$cnt = -1
			ForEach($tmp in $xIPAddress)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{ Data = "IP Address"; Value = $tmp; }
					$NicInformation += @{ Data = "Subnet Mask"; Value = $xIPSubnet[$cnt]; }
				}
			}
		}
		Else
		{
			$NicInformation += @{ Data = "IP Address"; Value = $xIPAddress; }
			$NicInformation += @{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }
			$NicInformation += @{ Data = "Subnet Mask"; Value = $xIPSubnet; }
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$NicInformation += @{ Data = "DHCP Enabled"; Value = $Nic.dhcpenabled; }
			$NicInformation += @{ Data = "DHCP Lease Obtained"; Value = $dhcpleaseobtaineddate; }
			$NicInformation += @{ Data = "DHCP Lease Expires"; Value = $dhcpleaseexpiresdate; }
			$NicInformation += @{ Data = "DHCP Server"; Value = $Nic.dhcpserver; }
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$NicInformation += @{ Data = "DNS Domain"; Value = $Nic.dnsdomain; }
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$NicInformation += @{ Data = "DNS Search Suffixes"; Value = $xnicdnsdomainsuffixsearchorder[0]; }
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$NicInformation += @{ Data = "DNS WINS Enabled"; Value = $xdnsenabledforwinsresolution; }
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			$NicInformation += @{ Data = "DNS Servers"; Value = $xnicdnsserversearchorder[0]; }
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$NicInformation += @{ Data = "NetBIOS Setting"; Value = $xTcpipNetbiosOptions; }
		$NicInformation += @{ Data = "WINS: Enabled LMHosts"; Value = $xwinsenablelmhostslookup; }
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$NicInformation += @{ Data = "Host Lookup File"; Value = $Nic.winshostlookupfile; }
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$NicInformation += @{ Data = "Primary Server"; Value = $Nic.winsprimaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$NicInformation += @{ Data = "Secondary Server"; Value = $Nic.winssecondaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$NicInformation += @{ Data = "Scope ID"; Value = $Nic.winsscopeid; }
		}
		$Table = AddWordTable -Hashtable $NicInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t: " $ThisNic.Name
		If($ThisNic.Name -ne $nic.description)
		{
			Line 2 "Description`t`t: " $nic.description
		}
		Line 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
		Line 2 "Manufacturer`t`t: " $ThisNic.manufacturer
		Line 2 "Availability`t`t: " $xAvailability
		Line 2 "Physical Address`t: " $nic.macaddress
		Line 2 "IP Address`t`t: " $xIPAddress[0]
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "" $tmp
			}
		}
		Line 2 "Default Gateway`t`t: " $Nic.Defaultipgateway
		Line 2 "Subnet Mask`t`t: " $xIPSubnet[0]
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "" $tmp
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			Line 2 "DHCP Enabled`t`t: " $nic.dhcpenabled
			Line 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
			Line 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
			Line 2 "DHCP Server`t`t:" $nic.dhcpserver
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			Line 2 "DNS Domain`t`t: " $nic.dnsdomain
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Search Suffixes`t:" $xnicdnsdomainsuffixsearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		Line 2 "DNS WINS Enabled`t: " $xdnsenabledforwinsresolution
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Servers`t`t:" $xnicdnsserversearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		Line 2 "NetBIOS Setting`t`t: " $xTcpipNetbiosOptions
		Line 2 "WINS:"
		Line 3 "Enabled LMHosts`t: " $xwinsenablelmhostslookup
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			Line 3 "Host Lookup File`t: " $nic.winshostlookupfile
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			Line 3 "Primary Server`t`t: " $nic.winsprimaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			Line 3 "Secondary Server`t: " $nic.winssecondaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			Line 3 "Scope ID`t`t: " $nic.winsscopeid
		}
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$ThisNic.Name,$htmlwhite)
		If($ThisNic.Name -ne $nic.description)
		{
			$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Nic.description,$htmlwhite))
		}
		$rowdata += @(,('Connection ID',($htmlsilver -bor $htmlbold),$ThisNic.NetConnectionID,$htmlwhite))
		$rowdata += @(,('Manufacturer',($htmlsilver -bor $htmlbold),$Nic.manufacturer,$htmlwhite))
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))
		$rowdata += @(,('Physical Address',($htmlsilver -bor $htmlbold),$Nic.macaddress,$htmlwhite))
		$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$xIPAddress[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		$rowdata += @(,('Default Gateway',($htmlsilver -bor $htmlbold),$Nic.Defaultipgateway,$htmlwhite))
		$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$xIPSubnet[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$rowdata += @(,('DHCP Enabled',($htmlsilver -bor $htmlbold),$Nic.dhcpenabled,$htmlwhite))
			$rowdata += @(,('DHCP Lease Obtained',($htmlsilver -bor $htmlbold),$dhcpleaseobtaineddate,$htmlwhite))
			$rowdata += @(,('DHCP Lease Expires',($htmlsilver -bor $htmlbold),$dhcpleaseexpiresdate,$htmlwhite))
			$rowdata += @(,('DHCP Server',($htmlsilver -bor $htmlbold),$Nic.dhcpserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$rowdata += @(,('DNS Domain',($htmlsilver -bor $htmlbold),$Nic.dnsdomain,$htmlwhite))
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Search Suffixes',($htmlsilver -bor $htmlbold),$xnicdnsdomainsuffixsearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('DNS WINS Enabled',($htmlsilver -bor $htmlbold),$xdnsenabledforwinsresolution,$htmlwhite))
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Servers',($htmlsilver -bor $htmlbold),$xnicdnsserversearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('NetBIOS Setting',($htmlsilver -bor $htmlbold),$xTcpipNetbiosOptions,$htmlwhite))
		$rowdata += @(,('WINS: Enabled LMHosts',($htmlsilver -bor $htmlbold),$xwinsenablelmhostslookup,$htmlwhite))
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$rowdata += @(,('Host Lookup File',($htmlsilver -bor $htmlbold),$Nic.winshostlookupfile,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$rowdata += @(,('Primary Server',($htmlsilver -bor $htmlbold),$Nic.winsprimaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$rowdata += @(,('Secondary Server',($htmlsilver -bor $htmlbold),$Nic.winssecondaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$rowdata += @(,('Scope ID',($htmlsilver -bor $htmlbold),$Nic.winsscopeid,$htmlwhite))
		}

		$msg = "Network Interface(s)"
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}
#endregion

Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. SMith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish

	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2' }

			'da-'	{ 'Automatisk tabel 2' }

			'de-'	{ 'Automatische Tabelle 2' }

			'en-'	{ 'Automatic Table 2' }

			'es-'	{ 'Tabla automática 2' }

			'fi-'	{ 'Automaattinen taulukko 2' }

			'fr-'	{ 'Sommaire Automatique 2' }

			'nb-'	{ 'Automatisk tabell 2' }

			'nl-'	{ 'Automatische inhoudsopgave 2' }

			'pt-'	{ 'Sumário Automático 2' }

			'sv-'	{ 'Automatisk innehållsförteckning2' }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("ViewMaster", "Secteur (foncé)", "Sémaphore",
					"Rétrospective", "Ion (foncé)", "Ion (clair)", "Intégrale",
					"Filigrane", "Facette", "Secteur (clair)", "À bandes", "Austin",
					"Guide", "Whisp", "Lignes latérales", "Quadrillage")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Mosaïques", "Ligne latérale", "Annuel", "Perspective",
					"Contraste", "Emplacements de bureau", "Moderne", "Blocs empilés",
					"Rayures fines", "Austère", "Transcendant", "Classique", "Quadrillage",
					"Exposition", "Alphabet", "Mots croisés", "Papier journal", "Austin", "Guide")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana",
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral",
						"Ion (Dark)", "Ion (Light)", "Motion", "Retrospect", "Semaphore",
						"Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

Function BuildPVSObject
{
	Param([string]$MCLIGetWhat = '', [string]$MCLIGetParameters = '', [string]$TextForErrorMsg = '')

	$error.Clear()

	If($MCLIGetParameters -ne '')
	{
		$MCLIGetResult = Mcli-Get "$($MCLIGetWhat)" -p "$($MCLIGetParameters)"
	}
	Else
	{
		$MCLIGetResult = Mcli-Get "$($MCLIGetWhat)"
	}

	If($error.Count -eq 0)
	{
		$PluralObject = @()
		$SingleObject = $Null
		ForEach($record in $MCLIGetResult)
		{
			If($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
			{
				If($SingleObject -ne $Null)
				{
					$PluralObject += $SingleObject
				}
				$SingleObject = new-object System.Object
			}

			$index = $record.IndexOf(':')
			If($index -gt 0)
			{
				$property = $record.SubString(0, $index)
				$value    = $record.SubString($index + 2)
				If($property -ne "Executing")
				{
					Add-Member -inputObject $SingleObject -MemberType NoteProperty -Name $property -Value $value
				}
			}
		}
		$PluralObject += $SingleObject
		Return $PluralObject
	}
	Else 
	{
		WriteWordLine 0 0 "$($TextForErrorMsg) could not be retrieved"
		WriteWordLine 0 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
	}
}

Function DeviceStatus
{
	Param($xDevice)

	If($xDevice -eq $Null -or $xDevice.status -eq "" -or $xDevice.status -eq "0")
	{
		WriteWordLine 0 3 "Target device inactive"
	}
	Else
	{
		WriteWordLine 0 3 "Target device active"
		WriteWordLine 0 3 "IP Address`t`t: " $xDevice.ip
		WriteWordLine 0 3 "Server`t`t`t: " -nonewline
		WriteWordLine 0 0 "$($xDevice.serverName) `($($xDevice.serverIpConnection)`: $($xDevice.serverPortConnection)`)"
		WriteWordLine 0 3 "Retries`t`t`t: " $xDevice.status
		WriteWordLine 0 3 "vDisk`t`t`t: " $xDevice.diskLocatorName
		WriteWordLine 0 3 "vDisk version`t`t: " $xDevice.diskVersion
		WriteWordLine 0 3 "vDisk name`t`t: " $xDevice.diskFileName
		WriteWordLine 0 3 "vDisk access`t`t: " -nonewline
		Switch ($xDevice.diskVersionAccess)
		{
			0 {WriteWordLine 0 0 "Production"}
			1 {WriteWordLine 0 0 "Test"}
			2 {WriteWordLine 0 0 "Maintenance"}
			3 {WriteWordLine 0 0 "Personal vDisk"}
			Default {WriteWordLine 0 0 "vDisk access type could not be determined: $($xDevice.diskVersionAccess)"}
		}
		If($PVSVersion -eq "7")
		{
			WriteWordLine 0 3 "Local write cache disk`t:$($xDevice.localWriteCacheDiskSize)GB"
			WriteWordLine 0 3 "Boot mode`t`t:" -nonewline
			Switch($xDevice.bdmBoot)
			{
				0 {WriteWordLine 0 0 "PXE boot"}
				1 {WriteWordLine 0 0 "BDM disk"}
				Default {WriteWordLine 0 0 "Boot mode could not be determined: $($xDevice.bdmBoot)"}
			}
		}
		Switch($xDevice.licenseType)
		{
			0 {WriteWordLine 0 3 "No License"}
			1 {WriteWordLine 0 3 "Desktop License"}
			2 {WriteWordLine 0 3 "Server License"}
			5 {WriteWordLine 0 3 "OEM SmartClient License"}
			6 {WriteWordLine 0 3 "XenApp License"}
			7 {WriteWordLine 0 3 "XenDesktop License"}
			Default {WriteWordLine 0 0 "Device license type could not be determined: $($xDevice.licenseType)"}
		}
		
		WriteWordLine 0 2 "Logging"
		WriteWordLine 0 3 "Logging level`t`t: " -nonewline
		Switch ($xDevice.logLevel)
		{
			0   {WriteWordLine 0 0 "Off"    }
			1   {WriteWordLine 0 0 "Fatal"  }
			2   {WriteWordLine 0 0 "Error"  }
			3   {WriteWordLine 0 0 "Warning"}
			4   {WriteWordLine 0 0 "Info"   }
			5   {WriteWordLine 0 0 "Debug"  }
			6   {WriteWordLine 0 0 "Trace"  }
			Default {WriteWordLine 0 0 "Logging level could not be determined: $($xDevice.logLevel)"}
		}
		
		WriteWordLine 0 0 ""
	}
}

Function SecondsToMinutes
{
	Param($xVal)
	
	If([int]$xVal -lt 60)
	{
		Return "0:$xVal"
	}
	$xMinutes = ([int]($xVal / 60)).ToString()
	$xSeconds = ([int]($xVal % 60)).ToString().PadLeft(2, "0")
	Return "$xMinutes`:$xSeconds"
}

Function Check-NeededPSSnapins
{
	Param([parameter(Mandatory = $True)][alias("Snapin")][string[]]$Snapins)

	#Function specifics
	$MissingSnapins = @()
	[bool]$FoundMissingSnapin = $False
	$LoadedSnapins = @()
	$RegisteredSnapins = @()

	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += get-pssnapin | % {$_.name}
	$registeredSnapins += get-pssnapin -Registered | % {$_.name}

	ForEach($Snapin in $Snapins)
	{
		#check if the snapin is loaded
		If(!($LoadedSnapins -like $snapin))
		{
			#Check if the snapin is missing
			If(!($RegisteredSnapins -like $Snapin))
			{
				#set the flag if it's not already
				If(!($FoundMissingSnapin))
				{
					$FoundMissingSnapin = $True
				}
				#add the entry to the list
				$MissingSnapins += $Snapin
			}
			Else
			{
				#Snapin is registered, but not loaded, loading it now:
				Write-Host "Loading Windows PowerShell snap-in: $snapin"
				Add-PSSnapin -Name $snapin -EA 0
			}
		}
	}

	If($FoundMissingSnapin)
	{
		Write-Warning "Missing Windows PowerShell snap-ins Detected:"
		$missingSnapins | % {Write-Warning "($_)"}
		return $False
	}
	Else
	{
		Return $True
	}
}

Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com
#for creating the formatted text report
#created March 2011
#updated March 2014
{
	Param( [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "`r`n", [switch]$nonewline )
	While( $tabs -gt 0 ) { $Global:Output += "`t"; $tabs--; }
	If( $nonewline )
	{
		$Global:Output += $name + $value
	}
	Else
	{
		$Global:Output += $name + $value + $newline
	}
}
	
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $myHash.Word_NoSpacing}
		1 {$Script:Selection.Style = $myHash.Word_Heading1}
		2 {$Script:Selection.Style = $myHash.Word_Heading2}
		3 {$Script:Selection.Style = $myHash.Word_Heading3}
		4 {$Script:Selection.Style = $myHash.Word_Heading4}
		Default {$Script:Selection.Style = $myHash.Word_NoSpacing}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop = $properties | ForEach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function AbortScript
{
	$Script:Word.quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>
Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Columns = $null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Headers = $null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$true)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Columns -eq $null) -and ($Headers -ne $null)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $null;
		}
		ElseIf(($Columns -ne $null) -and ($Headers -ne $null)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
        [System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Columns -eq $null) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Columns -eq $null) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{ 
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
                        [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    [ref] $null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $true);
			$ConvertToTableArguments.Add("ApplyShading", $true);
			$ConvertToTableArguments.Add("ApplyFont", $true);
			$ConvertToTableArguments.Add("ApplyColor", $true);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $true); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $true);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $true);
			$ConvertToTableArguments.Add("ApplyLastColumn", $true);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$null,                                          # Modifiers
			$null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		#the next line causes the heading row to flow across page breaks
		$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) returns a single Word COM cells object.
#>
Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end foreach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
				If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>
Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Company Name : $($Script:CoName)"
	Write-Verbose "$(Get-Date): Cover Page   : $($CoverPage)"
	Write-Verbose "$(Get-Date): User Name    : $($UserName)"
	Write-Verbose "$(Get-Date): Save As PDF  : $($PDF)"
	Write-Verbose "$(Get-Date): Save As TEXT : $($TEXT)"
	Write-Verbose "$(Get-Date): Save As WORD : $($MSWORD)"
	Write-Verbose "$(Get-Date): Save As HTML : $($HTML)"
	Write-Verbose "$(Get-Date): Add DateTime : $($AddDateTime)"
	Write-Verbose "$(Get-Date): HW Inventory : $($Hardware)"
	Write-Verbose "$(Get-Date): Filename1    : $($Script:FileName1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2    : $($Script:FileName2)"
	}
	Write-Verbose "$(Get-Date): OS Detected  : $($RunningOS)"
	Write-Verbose "$(Get-Date): PSUICulture  : $($PSUICulture)"
	Write-Verbose "$(Get-Date): PSCulture    : $($PSCulture)"
	Write-Verbose "$(Get-Date): Word version : $($Script:WordProduct)"
	Write-Verbose "$(Get-Date): Word language: $($Script:WordLanguageValue)"
	Write-Verbose "$(Get-Date): PoSH version : $($Host.Version)"
	Write-Verbose "$(Get-Date): PVS version  : $($PVSFullVersion)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	if( $object )
	{
		If( ( gm -Name $topLevel -InputObject $object ) )
		{
			If( ( gm -Name $secondLevel -InputObject $object.$topLevel ) )
			{
				Return $True
			}
		}
	}
	Return $False
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject.  Ignore the next message."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0
	
	If(!$? -or $Script:Word -eq $Null)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
		Exit
	}

	Write-Verbose "$(Get-Date): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tUnable to determine the Word language value.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}
	Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tMicrosoft Word 2007 is no longer supported.`n`n`t`tScript will end.`n`n"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($CoName))
	{
		Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "`n`n`t`tFor $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	ShowScriptOptions

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013
	$BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach{
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($BuildingBlocks -ne $Null)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($part -ne $Null)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Script:Doc -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Script:Selection -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $($myHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($myHash.Word_TableOfContents)
		If($toc -eq $Null)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($myHash.Word_TableOfContents) could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Company" $Script:CoName
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Title" $Script:title
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Author" $username

			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Subject" $SubjectTitle

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}

			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $Script:CoName"
			}

			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running Word 2010 and detected operating system $($RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running Word 2013 and detected operating system $($RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Deleting $($Script:FileName1) since only $($Script:FileName2) is needed"
		Remove-Item $Script:FileName1
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
}

Function SaveandCloseTextDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	Write-Output $Global:Output | Out-File $Script:Filename1
}

Function SaveandCloseHTMLDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	$pwdpath = $pwd.Path

	If($pwdpath.EndsWith("\"))
	{
		#remove the trailing \
		$pwdpath = $pwdpath.SubString(0, ($pwdpath.Length - 1))
	}

	#set $filename1 and $filename2 with no file extension
	If($AddDateTime)
	{
		[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName)"
		If($PDF)
		{
			[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName)"
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq

		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).docx"
			If($PDF)
			{
				[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName).pdf"
			}
		}

		SetupWord
	}
	ElseIf($Text)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).txt"
		}
	}
	ElseIf($HTML)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).html"
		}
	}
}

#script begins

$script:startTime = get-date

If($TEXT)
{
	$global:output = ""
}

Write-Verbose "$(Get-Date): Checking for McliPSSnapin"
If(!(Check-NeededPSSnapins "McliPSSnapIn")){
	#We're missing Citrix Snapins that we need
	$ErrorActionPreference = $SaveEAPreference
	Write-Error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Script will now close."
	Exit
}

#setup remoting if $AdminAddress is not empty
[bool]$Remoting = $False
If(![System.String]::IsNullOrEmpty($AdminAddress))
{
	If(![System.String]::IsNullOrEmpty($User))
	{
		If([System.String]::IsNullOrEmpty($Domain))
		{
			$Domain = Read-Host "Domain name for user is required.  Enter Domain name for user"
		}		

		If([System.String]::IsNullOrEmpty($Password))
		{
			$Password = Read-Host "Password for user is required.  Enter password for user"
		}		
		$error.Clear()
		mcli-run SetupConnection -p server="$($AdminAddress)",user="$($User)",domain="$($Domain)",password="$($Password)"
	}
	Else
	{
		$error.Clear()
		mcli-run SetupConnection -p server="$($AdminAddress)"
	}

	If($error.Count -eq 0)
	{
		$Remoting = $True
		Write-Verbose "$(Get-Date): This script is being run remotely against server $($AdminAddress)"
		If(![System.String]::IsNullOrEmpty($User))
		{
			Write-Verbose "$(Get-Date): User=$($User)"
			Write-Verbose "$(Get-Date): Domain=$($Domain)"
		}
	}
	Else 
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Warning "Remoting could not be setup to server $($AdminAddress)"
		Write-Warning "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
		Write-Warning "Script cannot continue"
		Exit
	}
}

Write-Verbose "$(Get-Date): Verifying PVS SOAP and Stream Services are running"
$soapserver = $Null
$StreamService = $Null

If($Remoting)
{
	$soapserver = Get-Service -ComputerName $AdminAddress -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Soap Server*"}
	$StreamService = Get-Service -ComputerName $AdminAddress -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Stream Service*"}
}
Else
{
	$soapserver = Get-Service -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Soap Server*"}
	$StreamService = Get-Service -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Stream Service*"}
}

If($soapserver.Status -ne "Running")
{
	$ErrorActionPreference = $SaveEAPreference
	If($Remoting)
	{
		Write-Warning "The Citrix PVS Soap Server service is not Started on server $($AdminAddress)"
	}
	Else
	{
		Write-Warning "The Citrix PVS Soap Server service is not Started"
	}
	Write-Error "Script cannot continue.  See message above."
	Exit
}

If($StreamService.Status -ne "Running")
{
	$ErrorActionPreference = $SaveEAPreference
	If($Remoting)
	{
		Write-Warning "The Citrix PVS Stream Service service is not Started on server $($AdminAddress)"
	}
	Else
	{
		Write-Warning "The Citrix PVS Stream Service service is not Started"
	}
	Write-Error "Script cannot continue.  See message above."
	Exit
}

#get PVS major version
Write-Verbose "$(Get-Date): Getting PVS version info"

$error.Clear()
$tempversion = mcli-info version
If($? -and $error.Count -eq 0)
{
	#build PVS version values
	$version = new-object System.Object 
	ForEach($record in $tempversion)
	{
		$index = $record.IndexOf(':')
		If($index -gt 0)
		{
			$property = $record.SubString(0, $index)
			$value = $record.SubString($index + 2)
			Add-Member -inputObject $version -MemberType NoteProperty -Name $property -Value $value
		}
	}
} 
Else 
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Warning "PVS version information could not be retrieved"
	[int]$NumErrors = $Error.Count
	For($x=0; $x -le $NumErrors; $x++)
	{
		Write-Warning "Error(s) returned: " $error[$x]
	}
	Write-Error "Script is terminating"
	#without version info, script should not proceed
	Exit
}

$PVSVersion     = $Version.mapiVersion.SubString(0,1)
$PVSFullVersion = $Version.mapiVersion.SubString(0,3)
[string]$tempversion    = $Null
[string]$version        = $Null
[bool]$FarmAutoAddEnabled = $False

#build PVS farm values
Write-Verbose "$(Get-Date): Build PVS farm values"
#there can only be one farm
[string]$GetWhat = "Farm"
[string]$GetParam = ""
[string]$ErrorTxt = "PVS Farm information"
$farm = BuildPVSObject $GetWhat $GetParam $ErrorTxt

If($Farm -eq $Null)
{
	#without farm info, script should not proceed
	$ErrorActionPreference = $SaveEAPreference
	Write-Error "PVS Farm information could not be retrieved.  Script is terminating."
	Exit
}

[string]$FarmName = $farm.FarmName
[string]$Title="Inventory Report for the $($FarmName) Farm"
SetFileName1andFileName2 "$($farm.FarmName)"

Write-Verbose "$(Get-Date): Processing PVS Farm Information"
$selection.InsertNewPage()
WriteWordLine 1 0 "PVS Farm Information"
#general tab
WriteWordLine 2 0 "General"
If(![String]::IsNullOrEmpty($farm.description))
{
	WriteWordLine 0 1 "Name`t`t: " $farm.farmName
	WriteWordLine 0 1 "Description`t: " $farm.description
}
Else
{
	WriteWordLine 0 1 "Name: " $farm.farmName
}

#security tab
Write-Verbose "$(Get-Date): `tProcessing Security Tab"
WriteWordLine 2 0 "Security"
WriteWordLine 0 1 "Groups with Farm Administrator access:"
#build security tab values
$GetWhat = "authgroup"
$GetParam = "farm = 1"
$ErrorTxt = "Groups with Farm Administrator access"
$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

If($AuthGroups -ne $Null)
{
	ForEach($Group in $authgroups)
	{
		If($Group.authGroupName)
		{
			WriteWordLine 0 2 $Group.authGroupName
		}
	}
}

#groups tab
Write-Verbose "$(Get-Date): `tProcessing Groups Tab"
WriteWordLine 2 0 "Groups"
WriteWordLine 0 1 "All the Security Groups that can be assigned access rights:"
$GetWhat = "authgroup"
$GetParam = ""
$ErrorTxt = "Security Groups information"
$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

If($AuthGroups -ne $Null)
{
	ForEach($Group in $authgroups)
	{
		If($Group.authGroupName)
		{
			WriteWordLine 0 2 $Group.authGroupName
		}
	}
}

#licensing tab
Write-Verbose "$(Get-Date): `tProcessing Licensing Tab"
WriteWordLine 2 0 "Licensing"
WriteWordLine 0 1 "License server name`t: " $farm.licenseServer
WriteWordLine 0 1 "License server port`t: " $farm.licenseServerPort
If($PVSVersion -eq "5")
{
	WriteWordLine 0 1 "Use Datacenter licenses for desktops if no Desktop licenses are available: " -nonewline
	If($farm.licenseTradeUp -eq "1")
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
}

#options tab
Write-Verbose "$(Get-Date): `tProcessing Options Tab"
WriteWordLine 2 0 "Options"
WriteWordLine 0 1 "Auto-Add"
WriteWordLine 0 2 "Enable auto-add: " -nonewline
If($farm.autoAddEnabled -eq "1")
{
	WriteWordLine 0 0 "Yes"
	WriteWordLine 0 3 "Add new devices to this site: " $farm.DefaultSiteName
	$FarmAutoAddEnabled = $True
}
Else
{
	WriteWordLine 0 0 "No"	
	$FarmAutoAddEnabled = $False
}
WriteWordLine 0 1 "Auditing"
WriteWordLine 0 2 "Enable auditing: " -nonewline
If($farm.auditingEnabled -eq "1")
{
	WriteWordLine 0 0 "Yes"
}
Else
{
	WriteWordLine 0 0 "No"
}
WriteWordLine 0 1 "Offline database support"
WriteWordLine 0 2 "Enable offline database support: " -nonewline
If($farm.offlineDatabaseSupportEnabled -eq "1")
{
	WriteWordLine 0 0 "Yes"	
}
Else
{
	WriteWordLine 0 0 "No"
}

If($PVSVersion -eq "6" -or $PVSVersion -eq "7")
{
	#vDisk Version tab
	Write-Verbose "$(Get-Date): `tProcessing vDisk Version Tab"
	WriteWordLine 2 0 "vDisk Version"
	WriteWordLine 0 1 "Alert if number of versions from base image exceeds`t`t: " $farm.maxVersions
	WriteWordLine 0 1 "Merge after automated vDisk update, if over alert threshold`t: " -nonewline
	If($farm.automaticMergeEnabled -eq "1")
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
	WriteWordLine 0 1 "Default access mode for new merge versions`t`t`t: " -nonewline
	Switch ($Farm.mergeMode)
	{
		0   {WriteWordLine 0 0 "Production" }
		1   {WriteWordLine 0 0 "Test"       }
		2   {WriteWordLine 0 0 "Maintenance"}
		Default {WriteWordLine 0 0 "Default access mode could not be determined: $($Farm.mergeMode)"}
	}
}

#status tab
Write-Verbose "$(Get-Date): `tProcessing Status Tab"
WriteWordLine 2 0 "Status"
WriteWordLine 0 1 "Current status of the farm:"
WriteWordLine 0 2 "Database server`t: " $farm.databaseServerName
If(![String]::IsNullOrEmpty($farm.databaseInstanceName))
{
	WriteWordLine 0 2 "Database instance`t: " $farm.databaseInstanceName
}
WriteWordLine 0 2 "Database`t`t: " $farm.databaseName
If(![String]::IsNullOrEmpty($farm.failoverPartnerServerName))
{
	WriteWordLine 0 2 "Failover Partner Server: " $farm.failoverPartnerServerName
}
If(![String]::IsNullOrEmpty($farm.failoverPartnerInstanceName))
{
	WriteWordLine 0 2 "Failover Partner Instance: " $farm.failoverPartnerInstanceName
}
If($Farm.adGroupsEnabled -eq "1")
{
	WriteWordLine 0 2 "Active Directory groups are used for access rights"
}
Else
{
	WriteWordLine 0 2 "Active Directory groups are not used for access rights"
}
Write-Verbose "$(Get-Date): "
	
$farm = $Null
$authgroups = $Null

#build site values
Write-Verbose "$(Get-Date): Processing Sites"
$AdvancedItems1 = @()
$AdvancedItems2 = @()
$GetWhat = "site"
$GetParam = ""
$ErrorTxt = "PVS Site information"
$PVSSites = BuildPVSObject $GetWhat $GetParam $ErrorTxt

ForEach($PVSSite in $PVSSites)
{
	Write-Verbose "$(Get-Date): `tProcessing Site $($PVSSite.siteName)"
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Site properties"
	#general tab
	WriteWordLine 2 0 "General"
	If(![String]::IsNullOrEmpty($PVSSite.description))
	{
		WriteWordLine 0 1 "Name`t`t: " $PVSSite.siteName
		WriteWordLine 0 1 "Description`t: " $PVSSite.description
	}
	Else
	{
		WriteWordLine 0 1 "Name: " $PVSSite.siteName
	}

	#security tab
	Write-Verbose "$(Get-Date): `t`tProcessing Security Tab"
	$temp = $PVSSite.SiteName
	$GetWhat = "authgroup"
	$GetParam = "sitename = $temp"
	$ErrorTxt = "Groups with Site Administrator access"
	$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	WriteWordLine 2 0 "Security"
	If($authGroups -ne $Null)
	{
		WriteWordLine 0 1 "Groups with Site Administrator access:"
		ForEach($Group in $authgroups)
		{
			WriteWordLine 0 2 $Group.authGroupName
		}
	}
	Else
	{
		WriteWordLine 0 1 "Groups with Site Administrator access: No Site Administrators defined"
	}

	#MAK tab
	#MAK User and Password are encrypted

	#options tab
	Write-Verbose "$(Get-Date): `t`tProcessing Options Tab"
	WriteWordLine 2 0 "Options"
	WriteWordLine 0 1 "Auto-Add"
	If($PVSVersion -eq "5" -or (($PVSVersion -eq "6" -or $PVSVersion -eq "7") -and $FarmAutoAddEnabled))
	{
		WriteWordLine 0 2 "Add new devices to this collection: " -nonewline
		If($PVSSite.DefaultCollectionName)
		{
			WriteWordLine 0 0 $PVSSite.DefaultCollectionName
		}
		Else
		{
			WriteWordLine 0 0 "<No Default collection>"
		}
	}
	If($PVSVersion -eq "6" -or $PVSVersion -eq "7")
	{
		If($PVSVersion -eq "6")
		{
			WriteWordLine 0 2 "Seconds between vDisk inventory scans: " $PVSSite.inventoryFilePollingInterval
		}

		#vDisk Update
		Write-Verbose "$(Get-Date): `t`tProcessing vDisk Update Tab"
		WriteWordLine 2 0 "vDisk Update"
		If($PVSSite.enableDiskUpdate -eq "1")
		{
			WriteWordLine 0 1 "Enable automatic vDisk updates on this site`t: " -nonewline
			WriteWordLine 0 0 "Yes"
			WriteWordLine 0 1 "Server to run vDisk updates for this site`t`t: " $PVSSite.diskUpdateServerName
		}
		Else
		{
			WriteWordLine 0 1 "Enable automatic vDisk updates on this site: No"
		}
	}

	#process all servers in site
	Write-Verbose "$(Get-Date): `t`tProcessing Servers in Site $($PVSSite.siteName)"
	$temp = $PVSSite.SiteName
	$GetWhat = "server"
	$GetParam = "sitename = $temp"
	$ErrorTxt = "Servers for Site $temp"
	$servers = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	WriteWordLine 2 0 "Servers"
	ForEach($Server in $Servers)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Server $($Server.serverName)"
		#general tab
		Write-Verbose "$(Get-Date): `t`t`t`tProcessing General Tab"
		WriteWordLine 3 0 $Server.serverName
		WriteWordLine 0 0 "Server Properties"
		WriteWordLine 0 1 "General"
		WriteWordLine 0 2 "Name`t`t: " $Server.serverName
		If(![String]::IsNullOrEmpty($Server.description))
		{
			WriteWordLine 0 2 "Description`t: " $Server.description
		}
		WriteWordLine 0 2 "Power Rating`t: " $Server.powerRating
		WriteWordLine 0 2 "Log events to the server's Windows Event Log: " -nonewline
		If($Server.eventLoggingEnabled -eq "1")
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
			
		Write-Verbose "$(Get-Date): `t`t`t`tProcessing Network Tab"
		WriteWordLine 0 1 "Network"
		If($PVSVersion -eq "7")
		{
			WriteWordLine 0 2 "Streaming IP addresses:"
		}
		Else
		{
			WriteWordLine 0 2 "IP addresses:"
		}
		$test = $Server.ip.ToString()
		$test1 = $test.replace(",","`n`t`t`t")
		WriteWordLine 0 3 $test1
		WriteWordLine 0 2 "Ports"
		WriteWordLine 0 3 "First port`t: " $Server.firstPort
		WriteWordLine 0 3 "Last port`t: " $Server.lastPort
		If($PVSVersion -eq "7")
		{
			WriteWordLine 0 2 "Management IP`t`t: " $Server.managementIp
		}
			
		Write-Verbose "$(Get-Date): `t`t`t`tProcessing Stores Tab"
		WriteWordLine 0 1 "Stores"
		#process all stores for this server
		Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Stores for server"
		$temp = $Server.serverName
		$GetWhat = "serverstore"
		$GetParam = "servername = $temp"
		$ErrorTxt = "Store information for server $temp"
		$stores = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		WriteWordLine 0 2 "Stores that this server supports:"

		If($Stores -ne $Null)
		{
			ForEach($store in $stores)
			{
				Write-Verbose "$(Get-Date): `t`t`t`t`t`tProcessing Store $($store.storename)"
				WriteWordLine 0 3 "Store`t: " $store.storename
				WriteWordLine 0 3 "Path`t: " -nonewline
				If($store.path.length -gt 0)
				{
					WriteWordLine 0 0 $store.path
				}
				Else
				{
					WriteWordLine 0 0 "<Using the Default path from the store>"
				}
				WriteWordLine 0 3 "Write cache paths: " -nonewline
				If($store.cachePath.length -gt 0)
				{
					WriteWordLine 0 0 $store.cachePath
				}
				Else
				{
					WriteWordLine 0 0 "<Using the Default path from the store>"
				}
				WriteWordLine 0 0 ""
			}
		}

		Write-Verbose "$(Get-Date): `t`t`t`tProcessing Options Tab"
		WriteWordLine 0 1 "Options"
		If($PVSVersion -eq "5")
		{
			WriteWordLine 0 2 "Enable automatic vDisk updates"
			WriteWordLine 0 3 "Check for new versions of a vDisk`t: " -nonewline
			If($Server.autoUpdateEnabled -eq "1")
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 3 "Check for incremental updates to a vDisk: " -nonewline
			If($Server.incrementalUpdateEnabled -eq "1")
			{
				WriteWordLine 0 0 "Yes"
				$AMorPM = "AM"
				$NumHour = [int]$Server.autoUpdateHour
				If($NumHour -ge 0 -and $NumHour -lt 12)
				{
					$AMorPM = "AM"
				}
				Else
				{
					$AMorPM = "PM"
				}
				If($NumHour -eq 0)
				{
					$NumHour +=  12
				}
				Else
				{
					$NumHour -=  12
				}
				$StrHour = [string]$NumHour
				If($StrHour.length -lt 2)
				{
					$StrHour = "0" + $StrHour
				}
				$tempMinute = ""
				If($Server.autoUpdateMinute.length -lt 2)
				{
					$tempMinute = "0" + $Server.autoUpdateMinute
				}
				WriteWordLine 0 3 "Check for updates daily at`t`t: $($StrHour)`:$($tempMinute) $($AMorPM)"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
		WriteWordLine 0 2 "Active directory"
		If($PVSVersion -eq "5")
		{
			WriteWordLine 0 3 "Enable automatic password support: " -nonewline
			If($Server.adMaxPasswordAgeEnabled -eq "1")
			{
				WriteWordLine 0 0 "Yes"
				WriteWordLine 0 3 "Change computer account password every $($Server.adMaxPasswordAge) days"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
		Else
		{
			WriteWordLine 0 3 "Automate computer account password updates`t: " -nonewline
			If($Server.adMaxPasswordAgeEnabled -eq "1")
			{
				WriteWordLine 0 0 "Yes"
				WriteWordLine 0 3 "Days between password updates`t`t: " $Server.adMaxPasswordAge
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
		
		If($PVSVersion -ne "7")
		{
			Write-Verbose "$(Get-Date): `t`t`t`tProcessing Logging Tab"
			WriteWordLine 0 1 "Logging"
			WriteWordLine 0 2 "Logging level: " -nonewline
			Switch ($Server.logLevel)
			{
				0   {WriteWordLine 0 0 "Off"    }
				1   {WriteWordLine 0 0 "Fatal"  }
				2   {WriteWordLine 0 0 "Error"  }
				3   {WriteWordLine 0 0 "Warning"}
				4   {WriteWordLine 0 0 "Info"   }
				5   {WriteWordLine 0 0 "Debug"  }
				6   {WriteWordLine 0 0 "Trace"  }
				Default {WriteWordLine 0 0 "Logging level could not be determined: $($Server.logLevel)"}
			}
			WriteWordLine 0 3 "File size maximum`t: $($Server.logFileSizeMax) (MB)"
			WriteWordLine 0 3 "Backup files maximum`t: " $Server.logFileBackupCopiesMax
			WriteWordLine 0 0 ""
		}
		
		#create array for appendix A
		
		Write-Verbose "$(Get-Date): `t`t`t`t`tGather Advanced server info for Appendix A and B"
		$obj1 = New-Object -TypeName PSObject
		$obj2 = New-Object -TypeName PSObject
		
		$obj1 | Add-Member -MemberType NoteProperty -Name ServerName              -Value $Server.serverName
		$obj1 | Add-Member -MemberType NoteProperty -Name ThreadsPerPort          -Value $Server.threadsPerPort
		$obj1 | Add-Member -MemberType NoteProperty -Name BuffersPerThread        -Value $Server.buffersPerThread
		$obj1 | Add-Member -MemberType NoteProperty -Name ServerCacheTimeout      -Value $Server.serverCacheTimeout
		$obj1 | Add-Member -MemberType NoteProperty -Name LocalConcurrentIOLimit  -Value $Server.localConcurrentIoLimit
		$obj1 | Add-Member -MemberType NoteProperty -Name RemoteConcurrentIOLimit -Value $Server.remoteConcurrentIoLimit
		$obj1 | Add-Member -MemberType NoteProperty -Name maxTransmissionUnits    -Value $Server.maxTransmissionUnits
		$obj1 | Add-Member -MemberType NoteProperty -Name IOBurstSize             -Value $Server.ioBurstSize
		$obj1 | Add-Member -MemberType NoteProperty -Name NonBlockingIOEnabled    -Value $Server.nonBlockingIoEnabled

		$obj2 | Add-Member -MemberType NoteProperty -Name ServerName              -Value $Server.serverName
		$obj2 | Add-Member -MemberType NoteProperty -Name BootPauseSeconds        -Value $Server.bootPauseSeconds
		$obj2 | Add-Member -MemberType NoteProperty -Name MaxBootSeconds          -Value $Server.maxBootSeconds
		$obj2 | Add-Member -MemberType NoteProperty -Name MaxBootDevicesAllowed   -Value $Server.maxBootDevicesAllowed
		$obj2 | Add-Member -MemberType NoteProperty -Name vDiskCreatePacing       -Value $Server.vDiskCreatePacing
		$obj2 | Add-Member -MemberType NoteProperty -Name LicenseTimeout          -Value $Server.licenseTimeout
		
		$AdvancedItems1 +=  $obj1
		$AdvancedItems2 +=  $obj2
		
		#advanced button at the bottom
		Write-Verbose "$(Get-Date): `t`t`t`tProcessing Server Advanced button"
		Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Server Tab"
		WriteWordLine 0 1 "Advanced"
		WriteWordLine 0 2 "Server"
		WriteWordLine 0 3 "Threads per port`t`t: " $Server.threadsPerPort
		WriteWordLine 0 3 "Buffers per thread`t`t: " $Server.buffersPerThread
		WriteWordLine 0 3 "Server cache timeout`t`t: $($Server.serverCacheTimeout) (seconds)"
		WriteWordLine 0 3 "Local concurrent I/O limit`t: $($Server.localConcurrentIoLimit) (transactions)"
		WriteWordLine 0 3 "Remote concurrent I/O limit`t: $($Server.remoteConcurrentIoLimit) (transactions)"

		Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Network Tab"
		WriteWordLine 0 2 "Network"
		WriteWordLine 0 3 "Ethernet MTU`t`t`t: $($Server.maxTransmissionUnits) (bytes)"
		WriteWordLine 0 3 "I/O burst size`t`t`t: $($Server.ioBurstSize) (KB)"
		WriteWordLine 0 3 "Enable non-blocking I/O for network communications: " -nonewline
		If($Server.nonBlockingIoEnabled -eq "1")
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}

		Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Pacing Tab"
		WriteWordLine 0 2 "Pacing"
		WriteWordLine 0 3 "Boot pause seconds`t`t: " $Server.bootPauseSeconds
		$MaxBootTime = SecondsToMinutes $Server.maxBootSeconds
		If($PVSVersion -eq "7")
		{
			WriteWordLine 0 3 "Maximum boot time`t`t: $($MaxBootTime) (minutes:seconds)"
		}
		Else
		{
			WriteWordLine 0 3 "Maximum boot time`t`t: $($MaxBootTime)"
		}
		WriteWordLine 0 3 "Maximum devices booting`t: $($Server.maxBootDevicesAllowed) devices"
		If($PVSVersion -eq "7")
		{
			WriteWordLine 0 3 "vDisk Creation pacing`t`t: $($Server.vDiskCreatePacing) milliseconds"
		}
		Else
		{
			WriteWordLine 0 3 "vDisk Creation pacing`t`t: " $Server.vDiskCreatePacing
		}

		Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Device Tab"
		WriteWordLine 0 2 "Device"
		$LicenseTimeout = SecondsToMinutes $Server.licenseTimeout
		If($PVSVersion -eq "7")
		{
			WriteWordLine 0 3 "License timeout`t`t`t: $($LicenseTimeout) (minutes:seconds)"
		}
		Else
		{
			WriteWordLine 0 3 "License timeout`t`t`t: $($LicenseTimeout)"
		}

		WriteWordLine 0 0 ""
		
		If($Hardware)
		{
			If(Test-Connection -ComputerName $server.servername -quiet -EA 0)
			{
				GetComputerWMIInfo $server.ServerName
			}
		}
	}

	#the properties for the servers have been processed. 
	#now to process the stuff available via a right-click on each server

	#Configure Bootstrap is first
	Write-Verbose "$(Get-Date): `t`t`tProcessing Bootstrap files"
	WriteWordLine 2 0 "Configure Bootstrap settings"
	ForEach($Server in $Servers)
	{
		Write-Verbose "$(Get-Date): `t`t`tTesting to see if $($server.ServerName) is online and reachable"
		If(Test-Connection -ComputerName $server.servername -quiet -EA 0)
		{
			Write-Verbose "$(Get-Date): `t`t`t`tProcessing Bootstrap files for Server $($server.servername)"
			#first get all bootstrap files for the server
			$temp = $server.serverName
			$GetWhat = "ServerBootstrapNames"
			$GetParam = "serverName = $temp"
			$ErrorTxt = "Server Bootstrap Name information"
			$BootstrapNames = BuildPVSObject $GetWhat $GetParam $ErrorTxt

			#Now that the list of bootstrap names has been gathered
			#We have the mandatory parameter to get the bootstrap info
			#there should be at least one bootstrap filename
			WriteWordLine 3 0 $Server.serverName
			If($Bootstrapnames -ne $Null)
			{
				#cannot use the BuildPVSObject Function here
				$serverbootstraps = @()
				ForEach($Bootstrapname in $Bootstrapnames)
				{
					#get serverbootstrap info
					$error.Clear()
					$tempserverbootstrap = Mcli-Get ServerBootstrap -p name="$($Bootstrapname.name)",servername="$($server.serverName)"
					If($error.Count -eq 0)
					{
						$serverbootstrap = $Null
						ForEach($record in $tempserverbootstrap)
						{
							If($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
							{
								If($serverbootstrap -ne $Null)
								{
									$serverbootstraps +=  $serverbootstrap
								}
								$serverbootstrap = new-object System.Object
								#add the bootstrapname name value to the serverbootstrap object
								$property = "BootstrapName"
								$value = $Bootstrapname.name
								Add-Member -inputObject $serverbootstrap -MemberType NoteProperty -Name $property -Value $value
							}
							$index = $record.IndexOf(':')
							If($index -gt 0)
							{
								$property = $record.SubString(0, $index)
								$value = $record.SubString($index + 2)
								If($property -ne "Executing")
								{
									Add-Member -inputObject $serverbootstrap -MemberType NoteProperty -Name $property -Value $value
								}
							}
						}
						$serverbootstraps +=  $serverbootstrap
					}
					Else
					{
						WriteWordLine 0 0 "Server Bootstrap information could not be retrieved"
						WriteWordLine 0 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
					}
				}
				If($ServerBootstraps -ne $Null)
				{
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Bootstrap file $($ServerBootstrap.Bootstrapname)"
					Write-Verbose "$(Get-Date): `t`t`t`t`t`tProcessing General Tab"
					WriteWordLine 0 1 "General"	
					ForEach($ServerBootstrap in $ServerBootstraps)
					{
						WriteWordLine 0 2 "Bootstrap file`t: " $ServerBootstrap.Bootstrapname
						If($ServerBootstrap.bootserver1_Ip -ne "0.0.0.0")
						{
							WriteWordLine 0 2 "IP Address`t: " $ServerBootstrap.bootserver1_Ip
							WriteWordLine 0 2 "Subnet Mask`t: " $ServerBootstrap.bootserver1_Netmask
							WriteWordLine 0 2 "Gateway`t: " $ServerBootstrap.bootserver1_Gateway
							WriteWordLine 0 2 "Port`t`t: " $ServerBootstrap.bootserver1_Port
						}
						If($ServerBootstrap.bootserver2_Ip -ne "0.0.0.0")
						{
							WriteWordLine 0 2 "IP Address`t: " $ServerBootstrap.bootserver2_Ip
							WriteWordLine 0 2 "Subnet Mask`t: " $ServerBootstrap.bootserver2_Netmask
							WriteWordLine 0 2 "Gateway`t: " $ServerBootstrap.bootserver2_Gateway
							WriteWordLine 0 2 "Port`t`t: " $ServerBootstrap.bootserver2_Port
						}
						If($ServerBootstrap.bootserver3_Ip -ne "0.0.0.0")
						{
							WriteWordLine 0 2 "IP Address`t: " $ServerBootstrap.bootserver3_Ip
							WriteWordLine 0 2 "Subnet Mask`t: " $ServerBootstrap.bootserver3_Netmask
							WriteWordLine 0 2 "Gateway`t: " $ServerBootstrap.bootserver3_Gateway
							WriteWordLine 0 2 "Port`t`t: " $ServerBootstrap.bootserver3_Port
						}
						If($ServerBootstrap.bootserver4_Ip -ne "0.0.0.0")
						{
							WriteWordLine 0 2 "IP Address`t: " $ServerBootstrap.bootserver4_Ip
							WriteWordLine 0 2 "Subnet Mask`t: " $ServerBootstrap.bootserver4_Netmask
							WriteWordLine 0 2 "Gateway`t: " $ServerBootstrap.bootserver4_Gateway
							WriteWordLine 0 2 "Port`t`t: " $ServerBootstrap.bootserver4_Port
						}
						WriteWordLine 0 0 ""
					}
				}
				Write-Verbose "$(Get-Date): `t`t`t`t`t`tProcessing Options Tab"
				WriteWordLine 0 1 "Options"
				WriteWordLine 0 2 "Verbose mode`t`t`t: " -nonewline
				If($ServerBootstrap.verboseMode -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 2 "Interrupt safe mode`t`t: " -nonewline
				If($ServerBootstrap.interruptSafeMode -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 2 "Advanced Memory Support`t: " -nonewline
				If($ServerBootstrap.paeMode -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 2 "Network recovery method`t: " -nonewline
				If($ServerBootstrap.bootFromHdOnFail -eq "0")
				{
					WriteWordLine 0 0 "Restore network connection"
				}
				Else
				{
					WriteWordLine 0 0 "Reboot to Hard Drive after $($ServerBootstrap.recoveryTime) seconds"
				}
				WriteWordLine 0 2 "Timeouts"
				WriteWordLine 0 3 "Login polling timeout`t: " -nonewline
				If($ServerBootstrap.pollingTimeout -eq "")
				{
					WriteWordLine 0 0 "5000 (milliseconds)"
				}
				Else
				{
					WriteWordLine 0 0 "$($ServerBootstrap.pollingTimeout) (milliseconds)"
				}
				WriteWordLine 0 3 "Login general timeout`t: " -nonewline
				If($ServerBootstrap.generalTimeout -eq "")
				{
					WriteWordLine 0 0 "5000 (milliseconds)"
				}
				Else
				{
					WriteWordLine 0 0 "$($ServerBootstrap.generalTimeout) (milliseconds)"
				}
			}
			Else
			{
				WriteWordLine 0 2 "No Bootstrap names available"
			}
		}
		Else
		{
			Write-Verbose "$(Get-Date): `t`t`t`tServer $($server.servername) is offline"
		}
	}		

	#process all vDisks in site
	Write-Verbose "$(Get-Date): `t`tProcessing all vDisks in site"
	$Temp = $PVSSite.SiteName
	$GetWhat = "DiskInfo"
	$GetParam = "siteName = $Temp"
	$ErrorTxt = "Disk information"
	$Disks = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	WriteWordLine 2 0 "vDisk Pool"
	If($Disks -ne $Null)
	{
		ForEach($Disk in $Disks)
		{
			Write-Verbose "$(Get-Date): `t`t`tProcessing vDisk $($Disk.diskLocatorName)"
			WriteWordLine 3 0 $Disk.diskLocatorName
			If($PVSVersion -eq "5")
			{
				#PVS 5.x
				WriteWordLine 0 1 "vDisk Properties"
				Write-Verbose "$(Get-Date): `t`t`t`tProcessing General Tab"
				WriteWordLine 0 2 "General"
				WriteWordLine 0 3 "Store`t`t`t: " $Disk.storeName
				WriteWordLine 0 3 "Site`t`t`t: " $Disk.siteName
				WriteWordLine 0 3 "Filename`t: " $Disk.diskLocatorName
				If(![String]::IsNullOrEmpty($Disk.description))
				{
					WriteWordLine 0 3 "Description`t`t: " $Disk.description
				}
				If(![String]::IsNullOrEmpty($Disk.menuText))
				{
					WriteWordLine 0 3 "BIOS menu text`t`t: " $Disk.menuText
				}
				If(![String]::IsNullOrEmpty($Disk.serverName))
				{
					WriteWordLine 0 3 "Use this server to provide the vDisk: " $Disk.serverName
				}
				Else
				{
					WriteWordLine 0 3 "Subnet Affinity`t`t: " -nonewline
					Switch ($Disk.subnetAffinity)
					{
						0 {WriteWordLine 0 0 "None"}
						1 {WriteWordLine 0 0 "Best Effort"}
						2 {WriteWordLine 0 0 "Fixed"}
						Default {WriteWordLine 0 0 "Subnet Affinity could not be determined: $($Disk.subnetAffinity)"}
					}
					WriteWordLine 0 3 "Rebalance Enabled`t: " -nonewline
					If($Disk.rebalanceEnabled -eq "1")
					{
						WriteWordLine 0 0 "Yes"
						WriteWordLine 0 3 "Trigger Percent`t`t: $($Disk.rebalanceTriggerPercent)"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
				}
				WriteWordLine 0 3 "Allow use of this vDisk`t: " -nonewline
				If($Disk.enabled -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}

				WriteWordLine 0 1 "vDisk File Properties"
				Write-Verbose "$(Get-Date): `t`t`t`tProcessing vDisk File Properties"
				Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing General Tab"
				WriteWordLine 0 2 "General"
				WriteWordLine 0 3 "Name`t`t: " $Disk.diskLocatorName
				WriteWordLine 0 3 "Size`t`t: " (($Disk.diskSize/1024)/1024)/1024 -nonewline
				WriteWordLine 0 0 " MB"
				If(![String]::IsNullOrEmpty($Disk.longDescription))
				{
					WriteWordLine 0 3 "Description`t: " $Disk.longDescription
				}
				If(![String]::IsNullOrEmpty($Disk.class))
				{
					WriteWordLine 0 3 "Class`t`t: " $Disk.class
				}
				If(![String]::IsNullOrEmpty($Disk.imageType))
				{
					WriteWordLine 0 3 "Type`t`t: " $Disk.imageType
				}

				Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Mode Tab"
				WriteWordLine 0 2 "Mode"
				WriteWordLine 0 3 "Access mode: " -nonewline
				If($Disk.writeCacheType -eq "0")
				{
					WriteWordLine 0 0 "Private Image (single device, read/write access)"
				}
				ElseIf($Disk.writeCacheType -eq "7")
				{
					WriteWordLine 0 0 "Difference Disk Image"
				}
				Else
				{
					WriteWordLine 0 0 "Standard Image (multi-device, read-only access)"
					WriteWordLine 0 3 "Cache type: " -nonewline
					Switch ($Disk.writeCacheType)
					{
						0   {WriteWordLine 0 0 "Private Image"}
						1   {WriteWordLine 0 0 "Cache on server"}
						2   {WriteWordLine 0 0 "Cache encrypted on server disk" }
						3   {
							WriteWordLine 0 0 "Cache in device RAM"
							WriteWordLine 0 3 "Cache Size: $($Disk.writeCacheSize) MBs"
							}
						4   {WriteWordLine 0 0 "Cache on device's HD"}
						5   {WriteWordLine 0 0 "Cache encrypted on device's hard disk"}
						6   {WriteWordLine 0 0 "RAM Disk"}
						7   {WriteWordLine 0 0 "Difference Disk"}
						Default {WriteWordLine 0 0 "Cache type could not be determined: $($Disk.writeCacheType)"}
					}
				}
				If($Disk.activationDateEnabled -eq "0")
				{
					WriteWordLine 0 3 "Enable automatic updates for the vDisk: " -nonewline
					If($Disk.autoUpdateEnabled -eq "1")
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					WriteWordLine 0 3 "Apply vDisk updates as soon as they are detected by the server"
				}
				Else
				{
					WriteWordLine 0 3 "Enable automatic updates for the vDisk`t`t: " -nonewline
					If($Disk.autoUpdateEnabled -eq "1")
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					WriteWordLine 0 3 "Schedule the next vDisk update to occur on`t: $($Disk.activeDate)"
				}
				Write-Verbose "$(Get-Date): `t`t`t`tProcessing Identification Tab"
				WriteWordLine 0 2 "Identification"
				WriteWordLine 0 3 "Version`t`t: Major:$($Disk.majorRelease) Minor:$($Disk.minorRelease) Build:$($Disk.build)"
				WriteWordLine 0 3 "Serial #`t`t: " $Disk.serialNumber
				If(![String]::IsNullOrEmpty($Disk.date))
				{
					WriteWordLine 0 3 "Date`t`t: " $Disk.date
				}
				If(![String]::IsNullOrEmpty($Disk.author))
				{
					WriteWordLine 0 3 "Author`t`t: " $Disk.author
				}
				If(![String]::IsNullOrEmpty($Disk.title))
				{
					WriteWordLine 0 3 "Title`t`t: " $Disk.title
				}
				If(![String]::IsNullOrEmpty($Disk.company))
				{
					WriteWordLine 0 3 "Company`t: " $Disk.company
				}
				If(![String]::IsNullOrEmpty($Disk.internalName))
				{
					If($Disk.internalName.Length -le 45)
					{
						WriteWordLine 0 3 "Internal name`t: " $Disk.internalName
					}
					Else
					{
						WriteWordLine 0 3 "Internal name`t:`n`t`t`t" $Disk.internalName
					}
				}
				If(![String]::IsNullOrEmpty($Disk.originalFile))
				{
					If($Disk.originalFile.Length -le 45)
					{
						WriteWordLine 0 3 "Original file`t: " $Disk.originalFile
					}
					Else
					{
						WriteWordLine 0 3 "Original file`t:`n`t`t`t" $Disk.originalFile
					}
				}
				If(![String]::IsNullOrEmpty($Disk.hardwareTarget))
				{
					WriteWordLine 0 3 "Hardware target: " $Disk.hardwareTarget
				}

				Write-Verbose "$(Get-Date): `t`t`t`tProcessing Volume Licensing Tab"
				WriteWordLine 0 2 "Microsoft Volume Licensing"
				WriteWordLine 0 3 "Microsoft license type: " -nonewline
				Switch ($Disk.licenseMode)
				{
					0 {WriteWordLine 0 0 "None"                          }
					1 {WriteWordLine 0 0 "Multiple Activation Key (MAK)" }
					2 {WriteWordLine 0 0 "Key Management Service (KMS)"  }
					Default {WriteWordLine 0 0 "Volume License Mode could not be determined: $($Disk.licenseMode)"}
				}
				#options tab
				Write-Verbose "$(Get-Date): `t`t`t`tProcessing Options Tab"
				WriteWordLine 0 2 "Options"
				WriteWordLine 0 3 "High availability (HA): " -nonewline
				If($Disk.haEnabled -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 3 "AD machine account password management: " -nonewline
				If($Disk.adPasswordEnabled -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				
				WriteWordLine 0 3 "Printer management: " -nonewline
				If($Disk.printerManagementEnabled -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				#end of PVS 5.x
			}
			Else
			{
				#PVS 6.x or 7.x
				WriteWordLine 0 1 "vDisk Properties"
				Write-Verbose "$(Get-Date): `t`t`tProcessing vDisk Properties"
				Write-Verbose "$(Get-Date): `t`t`t`tProcessing General Tab"
				WriteWordLine 0 2 "General"
				WriteWordLine 0 3 "Site`t`t: " $Disk.siteName
				WriteWordLine 0 3 "Store`t`t: " $Disk.storeName
				WriteWordLine 0 3 "Filename`t: " $Disk.diskLocatorName
				WriteWordLine 0 3 "Size`t`t: " (($Disk.diskSize/1024)/1024)/1024 -nonewline
				WriteWordLine 0 0 " MB"
				WriteWordLine 0 3 "VHD block size`t: " $Disk.vhdBlockSize -nonewline
				WriteWordLine 0 0 " KB"
				WriteWordLine 0 3 "Access mode`t: " -nonewline
				If($Disk.writeCacheType -eq "0")
				{
					WriteWordLine 0 0 "Private Image (single device, read/write access)"
				}
				Else
				{
					WriteWordLine 0 0 "Standard Image (multi-device, read-only access)"
					WriteWordLine 0 3 "Cache type`t: " -nonewline
					Switch ($Disk.writeCacheType)
					{
						0   {WriteWordLine 0 0 "Private Image"}
						1   {WriteWordLine 0 0 "Cache on server"}
						3   {
							WriteWordLine 0 0 "Cache in device RAM"
							WriteWordLine 0 3 "Cache Size: $($Disk.writeCacheSize) MBs"
							}
						4   {WriteWordLine 0 0 "Cache on device's hard disk"}
						6   {WriteWordLine 0 0 "RAM Disk"}
						7   {WriteWordLine 0 0 "Difference Disk"}
						9   {WriteWordLine 0 0 "Cache in device RAM with overflow on hard disk"}
						Default {WriteWordLine 0 0 "Cache type could not be determined: $($Disk.writeCacheType)"}
					}
				}
				If(![String]::IsNullOrEmpty($Disk.menuText))
				{
					WriteWordLine 0 3 "BIOS boot menu text`t`t`t: " $Disk.menuText
				}
				WriteWordLine 0 3 "Enable AD machine acct pwd mgmt`t: " -nonewline
				If($Disk.adPasswordEnabled -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				
				WriteWordLine 0 3 "Enable printer management`t`t: " -nonewline
				If($Disk.printerManagementEnabled -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 3 "Enable streaming of this vDisk`t`t: " -nonewline
				If($Disk.Enabled -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				
				Write-Verbose "$(Get-Date): `t`t`t`tProcessing Identification Tab"
				WriteWordLine 0 2 "Identification"
				If(![String]::IsNullOrEmpty($Disk.description))
				{
					WriteWordLine 0 3 "Description`t: " $Disk.description
				}
				If(![String]::IsNullOrEmpty($Disk.date))
				{
					WriteWordLine 0 3 "Date`t`t: " $Disk.date
				}
				If(![String]::IsNullOrEmpty($Disk.author))
				{
					WriteWordLine 0 3 "Author`t`t: " $Disk.author
				}
				If(![String]::IsNullOrEmpty($Disk.title))
				{
					WriteWordLine 0 3 "Title`t`t: " $Disk.title
				}
				If(![String]::IsNullOrEmpty($Disk.company))
				{
					WriteWordLine 0 3 "Company`t: " $Disk.company
				}
				If(![String]::IsNullOrEmpty($Disk.internalName))
				{
					If($Disk.internalName.Length -le 45)
					{
						WriteWordLine 0 3 "Internal name`t: " $Disk.internalName
					}
					Else
					{
						WriteWordLine 0 3 "Internal name`t:`n`t`t`t" $Disk.internalName
					}
				}
				If(![String]::IsNullOrEmpty($Disk.originalFile))
				{
					If($Disk.originalFile.Length -le 45)
					{
						WriteWordLine 0 3 "Original file`t: " $Disk.originalFile
					}
					Else
					{
						WriteWordLine 0 3 "Original file`t:`n`t`t`t" $Disk.originalFile
					}
				}
				If(![String]::IsNullOrEmpty($Disk.hardwareTarget))
				{
					WriteWordLine 0 3 "Hardware target: " $Disk.hardwareTarget
				}

				Write-Verbose "$(Get-Date): `t`t`t`tProcessing Volume Licensing Tab"
				WriteWordLine 0 2 "Microsoft Volume Licensing"
				WriteWordLine 0 3 "Microsoft license type: " -nonewline
				Switch ($Disk.licenseMode)
				{
					0 {WriteWordLine 0 0 "None"                          }
					1 {WriteWordLine 0 0 "Multiple Activation Key (MAK)" }
					2 {WriteWordLine 0 0 "Key Management Service (KMS)"  }
					Default {WriteWordLine 0 0 "Volume License Mode could not be determined: $($Disk.licenseMode)"}
				}

				Write-Verbose "$(Get-Date): `t`t`t`tProcessing Auto Update Tab"
				WriteWordLine 0 2 "Auto Update"
				If($Disk.activationDateEnabled -eq "0")
				{
					WriteWordLine 0 3 "Enable automatic updates for the vDisk: " -nonewline
					If($Disk.autoUpdateEnabled -eq "1")
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					WriteWordLine 0 3 "Apply vDisk updates as soon as they are detected by the server"
				}
				Else
				{
					WriteWordLine 0 3 "Enable automatic updates for the vDisk`t`t: " -nonewline
					If($Disk.autoUpdateEnabled -eq "1")
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					WriteWordLine 0 3 "Schedule the next vDisk update to occur on`t: $($Disk.activeDate)"
				}
				If(![String]::IsNullOrEmpty($Disk.class))
				{
					WriteWordLine 0 3 "Class`t: " $Disk.class
				}
				If(![String]::IsNullOrEmpty($Disk.imageType))
				{
					WriteWordLine 0 3 "Type`t: " $Disk.imageType
				}
				WriteWordLine 0 3 "Major #`t: " $Disk.majorRelease
				WriteWordLine 0 3 "Minor #`t: " $Disk.minorRelease
				WriteWordLine 0 3 "Build #`t: " $Disk.build
				WriteWordLine 0 3 "Serial #`t: " $Disk.serialNumber
				
				#process Versions menu
				#get versions info
				#thanks to the PVS Product team for their help in understanding the Versions information
				Write-Verbose "$(Get-Date): `t`t`tProcessing vDisk Versions"
				$VersionsObjects = @()
				$error.Clear()
				$MCLIGetResult = Mcli-Get DiskVersion -p diskLocatorName="$($Disk.diskLocatorName)",storeName="$($disk.storeName)",siteName="$($disk.siteName)"
				If($error.Count -eq 0)
				{
					#build versions object
					$PluralObject = @()
					$SingleObject = $Null
					ForEach($record in $MCLIGetResult)
					{
						If($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
						{
							If($SingleObject -ne $Null)
							{
								$PluralObject += $SingleObject
							}
							$SingleObject = new-object System.Object
						}

						$index = $record.IndexOf(':')
						If($index -gt 0)
						{
							$property = $record.SubString(0, $index)
							$value    = $record.SubString($index + 2)
							If($property -ne "Executing")
							{
								Add-Member -inputObject $SingleObject -MemberType NoteProperty -Name $property -Value $value
							}
						}
					}
					$PluralObject += $SingleObject
					$DiskVersions = $PluralObject
					
					If($DiskVersions -ne $Null)
					{
						WriteWordLine 0 1 "vDisk Versions"
						#get the current booting version
						#by default, the $DiskVersions object is in version number order lowest to highest
						#the initial or base version is 0 and always exists
						[string]$BootingVersion = "0"
						[bool]$BootOverride = $False
						ForEach($DiskVersion in $DiskVersions)
						{
							If($DiskVersion.access -eq "3")
							{
								#override i.e. manually selected boot version
								$BootingVersion = $DiskVersion.version
								$BootOverride = $True
								Break
							}
							ElseIf($DiskVersion.access -eq "0" -and $DiskVersion.IsPending -eq "0" )
							{
								$BootingVersion = $DiskVersion.version
								$BootOverride = $False
							}
						}
						
						WriteWordLine 0 2 "Boot production devices from version`t: " -NoNewLine
						If($BootOverride)
						{
							WriteWordLine 0 0 $BootingVersion
						}
						Else
						{
							WriteWordLine 0 0 "Newest released"
						}
						WriteWordLine 0 0 ""
						
						ForEach($DiskVersion in $DiskVersions)
						{
							Write-Verbose "$(Get-Date): `t`t`t`tProcessing vDisk Version $($DiskVersion.version)"
							WriteWordLine 0 2 "Version`t`t`t`t`t: " -NoNewLine
							If($DiskVersion.version -eq $BootingVersion)
							{
								WriteWordLine 0 0 "$($DiskVersion.version) (Current booting version)"
							}
							Else
							{
								WriteWordLine 0 0 $DiskVersion.version
							}
							WriteWordLine 0 2 "Created`t`t`t`t`t: " $DiskVersion.createDate
							If(![String]::IsNullOrEmpty($DiskVersion.scheduledDate))
							{
								WriteWordLine 0 2 "Released`t`t`t`t: " $DiskVersion.scheduledDate
							}
							WriteWordLine 0 2 "Devices`t`t`t`t`t: " $DiskVersion.deviceCount
							WriteWordLine 0 2 "Access`t`t`t`t`t: " -NoNewLine
							Switch ($DiskVersion.access)
							{
								"0" {WriteWordLine 0 0 "Production"}
								"1" {WriteWordLine 0 0 "Maintenance"}
								"2" {WriteWordLine 0 0 "Maintenance Highest Version"}
								"3" {WriteWordLine 0 0 "Override"}
								"4" {WriteWordLine 0 0 "Merge"}
								"5" {WriteWordLine 0 0 "Merge Maintenance"}
								"6" {WriteWordLine 0 0 "Merge Test"}
								"7" {WriteWordLine 0 0 "Test"}
								Default {WriteWordLine 0 0 "Access could not be determined: $($DiskVersion.access)"}
							}
							WriteWordLine 0 2 "Type`t`t`t`t`t: " -NoNewLine
							Switch ($DiskVersion.type)
							{
								"0" {WriteWordLine 0 0 "Base"}
								"1" {WriteWordLine 0 0 "Manual"}
								"2" {WriteWordLine 0 0 "Automatic"}
								"3" {WriteWordLine 0 0 "Merge"}
								"4" {WriteWordLine 0 0 "Merge Base"}
								Default {WriteWordLine 0 0 "Type could not be determined: $($DiskVersion.type)"}
							}
							If(![String]::IsNullOrEmpty($DiskVersion.description))
							{
								WriteWordLine 0 2 "Properties`t`t`t`t: " $DiskVersion.description
							}
							WriteWordLine 0 2 "Can Delete`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canDelete)
							{
								0 {WriteWordLine 0 0 "No"}
								1 {WriteWordLine 0 0 "Yes"}
							}
							WriteWordLine 0 2 "Can Merge`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canMerge)
							{
								0 {WriteWordLine 0 0 "No"}
								1 {WriteWordLine 0 0 "Yes"}
							}
							WriteWordLine 0 2 "Can Merge Base`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canMergeBase)
							{
								0 {WriteWordLine 0 0 "No"}
								1 {WriteWordLine 0 0 "Yes"}
							}
							WriteWordLine 0 2 "Can Promote`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canPromote)
							{
								0 {WriteWordLine 0 0 "No"}
								1 {WriteWordLine 0 0 "Yes"}
							}
							WriteWordLine 0 2 "Can Revert back to Test`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canRevertTest)
							{
								0 {WriteWordLine 0 0 "No"}
								1 {WriteWordLine 0 0 "Yes"}
							}
							WriteWordLine 0 2 "Can Revert back to Maintenance`t: "  -NoNewLine
							Switch ($DiskVersion.canRevertMaintenance)
							{
								0 {WriteWordLine 0 0 "No"}
								1 {WriteWordLine 0 0 "Yes"}
							}
							WriteWordLine 0 2 "Can Set Scheduled Date`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canSetScheduledDate)
							{
								0 {WriteWordLine 0 0 "No"}
								1 {WriteWordLine 0 0 "Yes"}
							}
							WriteWordLine 0 2 "Can Override`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canOverride)
							{
								0 {WriteWordLine 0 0 "No"}
								1 {WriteWordLine 0 0 "Yes"}
							}
							WriteWordLine 0 2 "Is Pending`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.isPending)
							{
								0 {WriteWordLine 0 0 "No, version Scheduled Date has occurred"}
								1 {WriteWordLine 0 0 "Yes, version Scheduled Date has not occurred"}
							}
							WriteWordLine 0 2 "Replication Status`t`t`t: " -NoNewLine
							Switch ($DiskVersion.goodInventoryStatus)
							{
								0 {WriteWordLine 0 0 "Not available on all servers"}
								1 {WriteWordLine 0 0 "Available on all servers"}
								Default {WriteWordLine 0 0 "Replication status could not be determined: $($DiskVersion.goodInventoryStatus)"}
							}
							WriteWordLine 0 2 "Disk Filename`t`t`t`t: " $DiskVersion.diskFileName
							WriteWordLine 0 0 ""
						}
					}
				}
				Else
				{
					WriteWordLine 0 0 "Disk Version information could not be retrieved"
					WriteWordLine 0 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
				}
				
				#process vDisk Load Balancing Menu
				Write-Verbose "$(Get-Date): `t`t`tProcessing vDisk Load Balancing Menu"
				WriteWordLine 3 1 "vDisk Load Balancing"
				If(![String]::IsNullOrEmpty($Disk.serverName))
				{
					WriteWordLine 0 2 "Use this server to provide the vDisk: " $Disk.serverName
				}
				Else
				{
					WriteWordLine 0 2 "Subnet Affinity`t`t: " -nonewline
					Switch ($Disk.subnetAffinity)
					{
						0 {WriteWordLine 0 0 "None"}
						1 {WriteWordLine 0 0 "Best Effort"}
						2 {WriteWordLine 0 0 "Fixed"}
						Default {WriteWordLine 0 0 "Subnet Affinity could not be determined: $($Disk.subnetAffinity)"}
					}
					WriteWordLine 0 2 "Rebalance Enabled`t: " -nonewline
					If($Disk.rebalanceEnabled -eq "1")
					{
						WriteWordLine 0 0 "Yes"
						WriteWordLine 0 2 "Trigger Percent`t`t: $($Disk.rebalanceTriggerPercent)"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
				}
			}#end of PVS 6.x
		}
	}

	#process all vDisk Update Management in site (PVS 6.x and 7 only)
	If($PVSVersion -eq "6" -or $PVSVersion -eq "7")
	{
		Write-Verbose "$(Get-Date): `t`tProcessing vDisk Update Management"
		$Temp = $PVSSite.SiteName
		$GetWhat = "UpdateTask"
		$GetParam = "siteName = $Temp"
		$ErrorTxt = "vDisk Update Management information"
		$Tasks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		WriteWordLine 2 0 " vDisk Update Management"
		If($Tasks -ne $Null)
		{
			If($PVSVersion -eq "6")
			{
				#process all virtual hosts for this site
				Write-Verbose "$(Get-Date): `t`t`tProcessing virtual hosts (PVS6)"
				WriteWordLine 0 1 "Hosts"
				$Temp = $PVSSite.SiteName
				$GetWhat = "VirtualHostingPool"
				$GetParam = "siteName = $Temp"
				$ErrorTxt = "Virtual Hosting Pool information"
				$vHosts = BuildPVSObject $GetWhat $GetParam $ErrorTxt
				If($vHosts -ne $Null)
				{
					WriteWordLine 3 0 "Hosts"
					ForEach($vHost in $vHosts)
					{
						Write-Verbose "$(Get-Date): `t`t`t`tProcessing virtual host $($vHost.virtualHostingPoolName)"
						Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing General Tab"
						WriteWordLine 4 0 $vHost.virtualHostingPoolName
						WriteWordLine 0 2 "General"
						WriteWordLine 0 3 "Type`t`t: " -nonewline
						Switch ($vHost.type)
						{
							0 {WriteWordLine 0 0 "Citrix XenServer"}
							1 {WriteWordLine 0 0 "Microsoft SCVMM/Hyper-V"}
							2 {WriteWordLine 0 0 "VMWare vSphere/ESX"}
							Default {WriteWordLine 0 0 "Virtualization Host type could not be determined: $($vHost.type)"}
						}
						WriteWordLine 0 3 "Name`t`t: " $vHost.virtualHostingPoolName
						If(![String]::IsNullOrEmpty($vHost.description))
						{
							WriteWordLine 0 3 "Description`t: " $vHost.description
						}
						WriteWordLine 0 3 "Host`t`t: " $vHost.server
						
						Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Advanced Tab"
						WriteWordLine 0 2 "Advanced"
						WriteWordLine 0 3 "Update limit`t`t: " $vHost.updateLimit
						WriteWordLine 0 3 "Update timeout`t`t: $($vHost.updateTimeout) minutes"
						WriteWordLine 0 3 "Shutdown timeout`t: $($vHost.shutdownTimeout) minutes"
						WriteWordLine 0 3 "Port`t`t`t: " $vHost.port
					}
				}
			}
			WriteWordLine 0 1 "vDisks"
			#process all the Update Managed vDisks for this site
			Write-Verbose "$(Get-Date): `t`t`tProcessing all Update Managed vDisks for this site"
			$Temp = $PVSSite.SiteName
			$GetParam = "siteName = $Temp"
			$GetWhat = "diskUpdateDevice"
			$ErrorTxt = "Update Managed vDisk information"
			$ManagedvDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			If($ManagedvDisks -ne $Null)
			{
				WriteWordLine 3 0 "vDisks"
				ForEach($ManagedvDisk in $ManagedvDisks)
				{
					Write-Verbose "$(Get-Date): `t`t`t`tProcessing Managed vDisk $($ManagedvDisk.storeName)`\$($ManagedvDisk.disklocatorName)"
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing General Tab"
					WriteWordLine 4 0 "$($ManagedvDisk.storeName)`\$($ManagedvDisk.disklocatorName)"
					WriteWordLine 0 2 "General"
					WriteWordLine 0 3 "vDisk`t`t: " "$($ManagedvDisk.storeName)`\$($ManagedvDisk.disklocatorName)"
					WriteWordLine 0 3 "Virtual Host Connection: " 
					WriteWordLine 0 4 $ManagedvDisk.virtualHostingPoolName
					WriteWordLine 0 3 "VM Name`t: " $ManagedvDisk.deviceName
					WriteWordLine 0 3 "VM MAC`t: " $ManagedvDisk.deviceMac
					WriteWordLine 0 3 "VM Port`t: " $ManagedvDisk.port
									
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Personality Tab"
					#process all personality strings for this device
					$Temp = $ManagedvDisk.deviceName
					$GetWhat = "DevicePersonality"
					$GetParam = "deviceName = $Temp"
					$ErrorTxt = "Device Personality Strings information"
					$PersonalityStrings = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($PersonalityStrings -ne $Null)
					{
						WriteWordLine 0 2 "Personality"
						ForEach($PersonalityString in $PersonalityStrings)
						{
							WriteWordLine 0 3 "Name: " $PersonalityString.Name
							WriteWordLine 0 3 "String: " $PersonalityString.Value
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Status Tab"
					WriteWordLine 0 2 "Status"
					$Temp = $ManagedvDisk.deviceId
					$GetWhat = "deviceInfo"
					$GetParam = "deviceId = $Temp"
					$ErrorTxt = "Device Info information"
					$Device = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					DeviceStatus $Device
					
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Logging Tab"
					WriteWordLine 0 2 "Logging"
					WriteWordLine 0 3 "Logging level: " -nonewline
					Switch ($ManagedvDisk.logLevel)
					{
						0   {WriteWordLine 0 0 "Off"    }
						1   {WriteWordLine 0 0 "Fatal"  }
						2   {WriteWordLine 0 0 "Error"  }
						3   {WriteWordLine 0 0 "Warning"}
						4   {WriteWordLine 0 0 "Info"   }
						5   {WriteWordLine 0 0 "Debug"  }
						6   {WriteWordLine 0 0 "Trace"  }
						Default {WriteWordLine 0 0 "Logging level could not be determined: $($Server.logLevel)"}
					}
				}
			}
			
			If($Tasks -ne $Null)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing all Update Managed Tasks for this site"
				ForEach($Task in $Tasks)
				{
					Write-Verbose "$(Get-Date): `t`t`t`tProcessing Task $($Task.updateTaskName)"
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing General Tab"
					WriteWordLine 0 1 "Tasks"
					WriteWordLine 0 2 "General"
					WriteWordLine 0 3 "Name`t`t: " $Task.updateTaskName
					If(![String]::IsNullOrEmpty($Task.description))
					{
						WriteWordLine 0 3 "Description`t: " $Task.description
					}
					WriteWordLine 0 3 "Disable this task: " -nonewline
					If($Task.enabled -eq "1")
					{
						WriteWordLine 0 0 "No"
					}
					Else
					{
						WriteWordLine 0 0 "Yes"
					}
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Schedule Tab"
					WriteWordLine 0 2 "Schedule"
					WriteWordLine 0 3 "Recurrence: " -nonewline
					Switch ($Task.recurrence)
					{
						0 {WriteWordLine 0 0 "None"}
						1 {WriteWordLine 0 0 "Daily Everyday"}
						2 {WriteWordLine 0 0 "Daily Weekdays only"}
						3 {WriteWordLine 0 0 "Weekly"}
						4 {WriteWordLine 0 0 "Monthly Date"}
						5 {WriteWordLine 0 0 "Monthly Type"}
						Default {WriteWordLine 0 0 "Recurrence type could not be determined: $($Task.recurrence)"}
					}
					If($Task.recurrence -ne "0")
					{
						$AMorPM = "AM"
						$NumHour = [int]$Task.Hour
						If($NumHour -ge 0 -and $NumHour -lt 12)
						{
							$AMorPM = "AM"
						}
						Else
						{
							$AMorPM = "PM"
						}
						If($NumHour -eq 0)
						{
							$NumHour += 12
						}
						Else
						{
							$NumHour -= 12
						}
						$StrHour = [string]$NumHour
						If($StrHour.length -lt 2)
						{
							$StrHour = "0" + $StrHour
						}
						$tempMinute = ""
						If($Task.Minute.length -lt 2)
						{
							$tempMinute = "0" + $Task.Minute
						}
						WriteWordLine 0 3 "Run the update at $($StrHour)`:$($tempMinute) $($AMorPM)"
					}
					If($Task.recurrence -eq "3")
					{
						$dayMask = @{
							1 = "Sunday"
							2 = "Monday";
							4 = "Tuesday";
							8 = "Wednesday";
							16 = "Thursday";
							32 = "Friday";
							64 = "Saturday"}
						For($i = 1; $i -le 128; $i = $i * 2)
						{
							If(($Task.dayMask -band $i) -ne 0)
							{
								WriteWordLine 0 4 $dayMask.$i
							}
						}
					}
					If($Task.recurrence -eq "4")
					{
						WriteWordLine 0 3 "On Date " $Task.date
					}
					If($Task.recurrence -eq "5")
					{
						WriteWordLine 0 3 "On " -nonewline
						Switch($Task.monthlyOffset)
						{
							1 {WriteWordLine 0 0 "First " -nonewline}
							2 {WriteWordLine 0 0 "Second " -nonewline}
							3 {WriteWordLine 0 0 "Third " -nonewline}
							4 {WriteWordLine 0 0 "Fourth " -nonewline}
							5 {WriteWordLine 0 0 "Last " -nonewline}
							Default {WriteWordLine 0 0 "Monthly Offset could not be determined: $($Task.monthlyOffset)"}
						}
						$dayMask = @{
							1 = "Sunday"
							2 = "Monday";
							4 = "Tuesday";
							8 = "Wednesday";
							16 = "Thursday";
							32 = "Friday";
							64 = "Saturday";
							128 = "Weekday"}
						For($i = 1; $i -le 128; $i = $i * 2)
						{
							If(($Task.dayMask -band $i) -ne 0)
							{
								WriteWordLine 0 0 $dayMask.$i
							}
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing vDisks Tab"
					WriteWordLine 0 2 "vDisks"
					WriteWordLine 0 3 "vDisks to be updated by this task:"
					$Temp = $ManagedvDisk.deviceId
					$GetWhat = "diskUpdateDevice"
					$GetParam = "deviceId = $Temp"
					$ErrorTxt = "Device Info information"
					$vDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($vDisks -ne $Null)
					{
						ForEach($vDisk in $vDisks)
						{
							WriteWordLine 0 4 "vDisk`t: " -nonewline
							WriteWordLine 0 0 "$($vDisk.storeName)`\$($vDisk.diskLocatorName)"
							WriteWordLine 0 4 "Host`t: " $vDisk.virtualHostingPoolName
							WriteWordLine 0 4 "VM`t: " $vDisk.deviceName
							WriteWordLine 0 0 ""
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing ESD Tab"
					WriteWordLine 0 2 "ESD"
					WriteWordLine 0 3 "ESD client to use: " -nonewline
					Switch($Task.esdType)
					{
						""     {WriteWordLine 0 0 "None (runs a custom script on the client)"}
						"WSUS" {WriteWordLine 0 0 "Microsoft Windows Update Service (WSUS)"}
						"SCCM" {WriteWordLine 0 0 "Microsoft System Center Configuration Manager (SCCM)"}
						Default {WriteWordLine 0 0 "ESD Client could not be determined: $($Task.esdType)"}
					}
					
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Scripts Tab"
					If(![String]::IsNullOrEmpty($Task.preUpdateScript) -or ![String]::IsNullOrEmpty($Task.preVmScript) -or ![String]::IsNullOrEmpty($Task.postVmScript) -or ![String]::IsNullOrEmpty($Task.postUpdateScript))
					{
						WriteWordLine 0 2 "Scripts"
						WriteWordLine 0 3 "Scripts that execute with the vDisk update processing:"
						If(![String]::IsNullOrEmpty($Task.preUpdateScript))
						{
							WriteWordLine 0 3 "Pre-update script`t: " $Task.preUpdateScript
						}
						If(![String]::IsNullOrEmpty($Task.preVmScript))
						{
							WriteWordLine 0 3 "Pre-startup script`t: " $Task.preVmScript
						}
						If(![String]::IsNullOrEmpty($Task.postVmScript))
						{
							WriteWordLine 0 3 "Post-shutdown script`t: " $Task.postVmScript
						}
						If(![String]::IsNullOrEmpty($Task.postUpdateScript))
						{
							WriteWordLine 0 3 "Post-update script`t: " $Task.postUpdateScript
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Access Tab"
					WriteWordLine 0 2 "Access"
					WriteWordLine 0 3 "Upon successful completion, access assigned to the vDisk: " -nonewline
					Switch($Task.postUpdateApprove)
					{
						0 {WriteWordLine 0 0 "Production"}
						1 {WriteWordLine 0 0 "Test"}
						2 {WriteWordLine 0 0 "Maintenance"}
						Default {WriteWordLine 0 0 "Access method for vDisk could not be determined: $($Task.postUpdateApprove)"}
					}
				}
			}
		}
	}

	#process all device collections in site
	Write-Verbose "$(Get-Date): `t`tProcessing all device collections in site"
	$Temp = $PVSSite.SiteName
	$GetWhat = "Collection"
	$GetParam = "siteName = $Temp"
	$ErrorTxt = "Device Collection information"
	$Collections = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	If($Collections -ne $Null)
	{
		WriteWordLine 2 0 "Device Collections"
		ForEach($Collection in $Collections)
		{
			Write-Verbose "$(Get-Date): `t`t`tProcessing Collection $($Collection.collectionName)"
			Write-Verbose "$(Get-Date): `t`t`t`tProcessing General Tab"
			WriteWordLine 3 0 $Collection.collectionName
			WriteWordLine 0 1 "General"
			If(![String]::IsNullOrEmpty($Collection.description))
			{
				WriteWordLine 0 2 "Name`t`t: " $Collection.collectionName
				WriteWordLine 0 2 "Description`t: " $Collection.description
			}
			Else
			{
				WriteWordLine 0 2 "Name: " $Collection.collectionName
			}

			Write-Verbose "$(Get-Date): `t`t`t`tProcessing Security Tab"
			WriteWordLine 0 2 "Security"
			$Temp = $Collection.collectionId
			$GetWhat = "authGroup"
			$GetParam = "collectionId = $Temp"
			$ErrorTxt = "Device Collection information"
			$AuthGroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

			$DeviceAdmins = $False
			If($AuthGroups -ne $Null)
			{
				WriteWordLine 0 3 "Groups with 'Device Administrator' access:"
				ForEach($AuthGroup in $AuthGroups)
				{
					$Temp = $authgroup.authGroupName
					$GetWhat = "authgroupusage"
					$GetParam = "authgroupname = $Temp"
					$ErrorTxt = "Device Collection Administrator usage information"
					$AuthGroupUsages = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($AuthGroupUsages -ne $Null)
					{
						ForEach($AuthGroupUsage in $AuthGroupUsages)
						{
							If($AuthGroupUsage.role -eq "300")
							{
								$DeviceAdmins = $True
								WriteWordLine 0 3 $authgroup.authGroupName
							}
						}
					}
				}
			}
			If(!$DeviceAdmins)
			{
				WriteWordLine 0 3 "Groups with 'Device Administrator' access`t: None defined"
			}

			$DeviceOperators = $False
			If($AuthGroups -ne $Null)
			{
				WriteWordLine 0 3 "Groups with 'Device Operator' access:"
				ForEach($AuthGroup in $AuthGroups)
				{
					$Temp = $authgroup.authGroupName
					$GetWhat = "authgroupusage"
					$GetParam = "authgroupname = $Temp"
					$ErrorTxt = "Device Collection Operator usage information"
					$AuthGroupUsages = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($AuthGroupUsages -ne $Null)
					{
						ForEach($AuthGroupUsage in $AuthGroupUsages)
						{
							If($AuthGroupUsage.role -eq "400")
							{
								$DeviceOperators = $True
								WriteWordLine 0 3 $authgroup.authGroupName
							}
						}
					}
				}
			}
			If(!$DeviceOperators)
			{
				WriteWordLine 0 3 "Groups with 'Device Operator' access`t`t: None defined"
			}

			Write-Verbose "$(Get-Date): `t`t`t`tProcessing Auto-Add Tab"
			WriteWordLine 0 2 "Auto-Add"
			If($FarmAutoAddEnabled)
			{
				WriteWordLine 0 3 "Template target device: " $Collection.templateDeviceName
				If(![String]::IsNullOrEmpty($Collection.autoAddPrefix) -or ![String]::IsNullOrEmpty($Collection.autoAddPrefix))
				{
					WriteWordLine 0 3 "Device Name"
				}
				If(![String]::IsNullOrEmpty($Collection.autoAddPrefix))
				{
					WriteWordLine 0 4 "Prefix`t`t`t: " $Collection.autoAddPrefix
				}
				WriteWordLine 0 4 "Length`t`t`t: " $Collection.autoAddNumberLength
				WriteWordLine 0 4 "Zero fill`t`t`t: " -nonewline
				If($Collection.autoAddZeroFill -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				If(![String]::IsNullOrEmpty($Collection.autoAddPrefix))
				{
					WriteWordLine 0 4 "Suffix`t`t`t: " $Collection.autoAddSuffix
				}
				WriteWordLine 0 4 "Last incremental #`t: " $Collection.lastAutoAddDeviceNumber
			}
			Else
			{
				WriteWordLine 0 3 "The auto-add feature is not enabled at the PVS Farm level"
			}
			#for each collection process each device
			Write-Verbose "$(Get-Date): `t`t`tProcessing each collection process for each device"
			$Temp = $Collection.collectionId
			$GetWhat = "deviceInfo"
			$GetParam = "collectionId = $Temp"
			$ErrorTxt = "Device Info information"
			$Devices = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			
			If($Devices -ne $Null)
			{
				ForEach($Device in $Devices)
				{
					Write-Verbose "$(Get-Date): `t`t`t`tProcessing Device $($Device.deviceName)"
					If($Device.type -eq "3")
					{
						WriteWordLine 0 1 "Device with Personal vDisk Properties"
					}
					Else
					{
						WriteWordLine 0 1 "Target Device Properties"
					}
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing General Tab"
					WriteWordLine 0 2 "General"
					WriteWordLine 0 3 "Name`t`t`t: " $Device.deviceName
					If(![String]::IsNullOrEmpty($Device.description))
					{
						WriteWordLine 0 3 "Description`t`t: " $Device.description
					}
					If(($PVSVersion -eq "6" -or $PVSVersion -eq "7") -and $Device.type -ne "3")
					{
						WriteWordLine 0 3 "Type`t`t`t: " -nonewline
						Switch ($Device.type)
						{
							0 {WriteWordLine 0 0 "Production"}
							1 {WriteWordLine 0 0 "Test"}
							2 {WriteWordLine 0 0 "Maintenance"}
							3 {WriteWordLine 0 0 "Personal vDisk"}
							Default {WriteWordLine 0 0 "Device type could not be determined: $($Device.type)"}
						}
					}
					If($Device.type -ne "3")
					{
						WriteWordLine 0 3 "Boot from`t`t: " -nonewline
						Switch ($Device.bootFrom)
						{
							1 {WriteWordLine 0 0 "vDisk"}
							2 {WriteWordLine 0 0 "Hard Disk"}
							3 {WriteWordLine 0 0 "Floppy Disk"}
							Default {WriteWordLine 0 0 "Boot from could not be determined: $($Device.bootFrom)"}
						}
					}
					WriteWordLine 0 3 "MAC`t`t`t: " $Device.deviceMac
					WriteWordLine 0 3 "Port`t`t`t: " $Device.port
					If($Device.type -ne "3")
					{
						WriteWordLine 0 3 "Class`t`t`t: " $Device.className
						WriteWordLine 0 3 "Disable this device`t: " -nonewline
						If($Device.enabled -eq "1")
						{
							WriteWordLine 0 0 "Unchecked"
						}
						Else
						{
							WriteWordLine 0 0 "Checked"
						}
					}
					Else
					{
						WriteWordLine 0 3 "vDisk`t`t`t: " $Device.diskLocatorName
						WriteWordLine 0 3 "Personal vDisk Drive`t: " $Device.pvdDriveLetter
					}
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing vDisks Tab"
					WriteWordLine 0 2 "vDisks"
					#process all vdisks for this device
					$Temp = $Device.deviceName
					$GetWhat = "DiskInfo"
					$GetParam = "deviceName = $Temp"
					$ErrorTxt = "Device vDisk information"
					$vDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($vDisks -ne $Null)
					{
						ForEach($vDisk in $vDisks)
						{
							WriteWordLine 0 3 "Name: " -nonewline
							WriteWordLine 0 0 "$($vDisk.storeName)`\$($vDisk.diskLocatorName)"
						}
					}
					WriteWordLine 0 3 "Options"
					WriteWordLine 0 4 "List local hard drive in boot menu: " -nonewline
					If($Device.localDiskEnabled -eq "1")
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					#process all bootstrap files for this device
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing all bootstrap files for this device"
					$Temp = $Device.deviceName
					$GetWhat = "DeviceBootstraps"
					$GetParam = "deviceName = $Temp"
					$ErrorTxt = "Device Bootstrap information"
					$Bootstraps = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($Bootstraps -ne $Null)
					{
						ForEach($Bootstrap in $Bootstraps)
						{
							WriteWordLine 0 4 "Custom bootstrap file: " -nonewline
							WriteWordLine 0 0 "$($Bootstrap.bootstrap) `($($Bootstrap.menuText)`)"
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Authentication Tab"
					WriteWordLine 0 2 "Authentication"
					WriteWordLine 0 3 "Type of authentication to use for this device: " -nonewline
					Switch($Device.authentication)
					{
						0 {WriteWordLine 0 0 "None"}
						1 {WriteWordLine 0 0 "Username and password"; WriteWordLine 0 4 "Username: " $Device.user; WriteWordLine 0 4 "Password: " $Device.password}
						2 {WriteWordLine 0 0 "External verification (User supplied method)"}
						Default {WriteWordLine 0 0 "Authentication type could not be determined: $($Device.authentication)"}
					}
					
					Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing Personality Tab"
					#process all personality strings for this device
					$Temp = $Device.deviceName
					$GetWhat = "DevicePersonality"
					$GetParam = "deviceName = $Temp"
					$ErrorTxt = "Device Personality Strings information"
					$PersonalityStrings = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($PersonalityStrings -ne $Null)
					{
						WriteWordLine 0 2 "Personality"
						ForEach($PersonalityString in $PersonalityStrings)
						{
							WriteWordLine 0 3 "Name: " $PersonalityString.Name
							WriteWordLine 0 3 "String: " $PersonalityString.Value
						}
					}
					
					WriteWordLine 0 2 "Status"
					DeviceStatus $Device
				}
			}
		}
	}

	#process all user groups in site (PVS 5.6 only)
	If($PVSVersion -eq "5")
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing all user groups in site"
		$Temp = $PVSSite.siteName
		$GetWhat = "UserGroup"
		$GetParam = "siteName = $Temp"
		$ErrorTxt = "User Group information"
		$UserGroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		WriteWordLine 2 0 "User Group Properties"
		If($UserGroups -ne $Null)
		{
			ForEach($UserGroup in $UserGroups)
			{
				Write-Verbose "$(Get-Date): `t`t`t`tProcessing User Group $($UserGroup.userGroupName)"
				Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing General Tab"
				WriteWordLine 0 1 "General"
				WriteWordLine 0 2 "Name`t`t`t: " $UserGroup.userGroupName
				If(![String]::IsNullOrEmpty($UserGroup.description))
				{
					WriteWordLine 0 2 "Description`t`t: " $UserGroup.description
				}
				If(![String]::IsNullOrEmpty($UserGroup.className))
				{
					WriteWordLine 0 2 "Class`t`t`t: " $UserGroup.className
				}
				WriteWordLine 0 2 "Disable this group`t: " -nonewline
				If($UserGroup.enabled -eq "1")
				{
					WriteWordLine 0 0 "No"
				}
				Else
				{
					WriteWordLine 0 0 "Yes"
				}
				#process all vDisks for usergroup
				Write-Verbose "$(Get-Date): Process all vDisks for user group"
				$Temp = $UserGroup.userGroupId
				$GetWhat = "DiskInfo"
				$GetParam = "userGroupId = $Temp"
				$ErrorTxt = "User Group Disk information"
				$vDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt

				Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing vDisk Tab"
				WriteWordLine 0 1 "vDisk"
				WriteWordLine 0 2 "vDisks for this user group:"
				If($vDisks -ne $Null)
				{
					ForEach($vDisk in $vDisks)
					{
						WriteWordLine 0 3 "$($vDisk.storeName)`\$($vDisk.diskLocatorName)"
					}
				}
			}
		}
	}
	
	#process all site views in site
	Write-Verbose "$(Get-Date): `t`tProcessing all site views in site"
	$Temp = $PVSSite.siteName
	$GetWhat = "SiteView"
	$GetParam = "siteName = $Temp"
	$ErrorTxt = "Site View information"
	$SiteViews = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	
	WriteWordLine 2 0 "Site Views"
	If($SiteViews -ne $Null)
	{
		ForEach($SiteView in $SiteViews)
		{
			Write-Verbose "$(Get-Date): `t`t`tProcessing Site View $($SiteView.siteViewName)"
			Write-Verbose "$(Get-Date): `t`t`t`tProcessing General Tab"
			WriteWordLine 3 0 $SiteView.siteViewName
			WriteWordLine 0 1 "View Properties"
			WriteWordLine 0 2 "General"
			If(![String]::IsNullOrEmpty($SiteView.description))
			{
				WriteWordLine 0 3 "Name`t`t: " $SiteView.siteViewName
				WriteWordLine 0 3 "Description`t: " $SiteView.description
			}
			Else
			{
				WriteWordLine 0 3 "Name: " $SiteView.siteViewName
			}
			
			Write-Verbose "$(Get-Date): `t`t`t`tProcessing Members Tab"
			WriteWordLine 0 2 "Members"
			#process each target device contained in the site view
			$Temp = $SiteView.siteViewId
			$GetWhat = "Device"
			$GetParam = "siteViewId = $Temp"
			$ErrorTxt = "Site View Device Members information"
			$Members = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			If($Members -ne $Null)
			{
				ForEach($Member in $Members)
				{
					WriteWordLine 0 3 $Member.deviceName
				}
			}
		}
	}
	Else
	{
		WriteWordLine 0 1 "There are no Site Views configured"
	}
	If($PVSVersion -eq "7")
	{
		#process all virtual hosts for this site
		Write-Verbose "$(Get-Date): `t`t`tProcessing virtual hosts (PVS7)"
		$Temp = $PVSSite.SiteName
		$GetWhat = "VirtualHostingPool"
		$GetParam = "siteName = $Temp"
		$ErrorTxt = "Virtual Hosting Pool information"
		$vHosts = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		If($vHosts -ne $Null)
		{
			WriteWordLine 2 0 "Hosts"
			ForEach($vHost in $vHosts)
			{
				Write-Verbose "$(Get-Date): `t`t`t`tProcessing virtual host $($vHost.virtualHostingPoolName)"
				Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing General Tab"
				WriteWordLine 4 0 $vHost.virtualHostingPoolName
				WriteWordLine 0 2 "General"
				WriteWordLine 0 3 "Type`t`t: " -nonewline
				Switch ($vHost.type)
				{
					0 {WriteWordLine 0 0 "Citrix XenServer"}
					1 {WriteWordLine 0 0 "Microsoft SCVMM/Hyper-V"}
					2 {WriteWordLine 0 0 "VMWare vSphere/ESX"}
					Default {WriteWordLine 0 0 "Virtualization Host type could not be determined: $($vHost.type)"}
				}
				WriteWordLine 0 3 "Name`t`t: " $vHost.virtualHostingPoolName
				If(![String]::IsNullOrEmpty($vHost.description))
				{
					WriteWordLine 0 3 "Description`t: " $vHost.description
				}
				WriteWordLine 0 3 "Host`t`t: " $vHost.server
				
				Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing vDisk Update Tab"
				WriteWordLine 0 2 "Update limit`t`t: " $vHost.updateLimit
				WriteWordLine 0 2 "Update timeout`t`t: $($vHost.updateTimeout) minutes"
				WriteWordLine 0 2 "Shutdown timeout`t: $($vHost.shutdownTimeout) minutes"
			}
			WriteWordLine 0 0 ""
		}
	}
	
	#add Audit Trail
	Write-Verbose "$(Get-Date): `t`t`tProcessing Audit Trail"
	$AuditTrailObjects = @()
	$error.Clear()
	
	#the audittrail call requires the dates in YYYY/MM/DD format
	$Sdate = '{0:yyyy/MM/dd}' -f $StartDate
	$Edate = '{0:yyyy/MM/dd}' -f $EndDate
	$MCLIGetResult = Mcli-Get AuditTrail -p siteName="$($PVSSite.siteName)",beginDate="$($SDate)",endDate="$($EDate)"
	If($error.Count -eq 0)
	{
		#build audit trail object
		$PluralObject = @()
		$SingleObject = $Null
		ForEach($record in $MCLIGetResult)
		{
			If($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
			{
				If($SingleObject -ne $Null)
				{
					$PluralObject += $SingleObject
				}
				$SingleObject = new-object System.Object
			}

			$index = $record.IndexOf(':')
			If($index -gt 0)
			{
				$property = $record.SubString(0, $index)
				$value    = $record.SubString($index + 2)
				If($property -ne "Executing")
				{
					Add-Member -inputObject $SingleObject -MemberType NoteProperty -Name $property -Value $value
				}
			}
		}
		$PluralObject += $SingleObject
		$Audits = $PluralObject
		
		If($Audits -ne $Null)
		{
			If($Audits -is [array])
			{
				[int]$NumAudits = $Audits.Count +1
			}
			Else
			{
				[int]$NumAudits = 2
			}
			$selection.InsertNewPage()
			WriteWordLine 2 0 "Audit Trail"
			WriteWordLine 0 0 "Audit Trail for dates $($StartDate) through $($EndDate)"
			Write-Verbose "$(Get-Date): `t`t$NumAudits Audit Trail entries found"

			If($MSWord -or $PDF)
			{
				WriteWordLine 0 1 "Services ($NumServices Services found)"

				## IB - replacement Services table generation utilising AddWordTable function

				## Create an array of hashtables to store our services
				[System.Collections.Hashtable[]] $AuditWordTable = @();
				## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
				[System.Collections.Hashtable[]] $HighlightedCells = @();
				## Seed the $Services row index from the second row
				[int] $CurrentServiceIndex = 2;
			}
			ElseIf($Text)
			{
			}
			ElseIf($HTML)
			{
			}
			
			ForEach($Audit in $Audits)
			{
				$Tmp = ""
				Switch([int]$Audit.action)
				{
					1 { $Tmp = "AddAuthGroup"}
					2 { $Tmp = "AddCollection"}
					3 { $Tmp = "AddDevice"}
					4 { $Tmp = "AddDiskLocator"}
					5 { $Tmp = "AddFarmView"}
					6 { $Tmp = "AddServer"}
					7 { $Tmp = "AddSite"}
					8 { $Tmp = "AddSiteView"}
					9 { $Tmp = "AddStore"}
					10 { $Tmp = "AddUserGroup"}
					11 { $Tmp = "AddVirtualHostingPool"}
					12 { $Tmp = "AddUpdateTask"}
					13 { $Tmp = "AddDiskUpdateDevice"}
					1001 { $Tmp = "DeleteAuthGroup"}
					1002 { $Tmp = "DeleteCollection"}
					1003 { $Tmp = "DeleteDevice"}
					1004 { $Tmp = "DeleteDeviceDiskCacheFile"}
					1005 { $Tmp = "DeleteDiskLocator"}
					1006 { $Tmp = "DeleteFarmView"}
					1007 { $Tmp = "DeleteServer"}
					1008 { $Tmp = "DeleteServerStore"}
					1009 { $Tmp = "DeleteSite"}
					1010 { $Tmp = "DeleteSiteView"}
					1011 { $Tmp = "DeleteStore"}
					1012 { $Tmp = "DeleteUserGroup"}
					1013 { $Tmp = "DeleteVirtualHostingPool"}
					1014 { $Tmp = "DeleteUpdateTask"}
					1015 { $Tmp = "DeleteDiskUpdateDevice"}
					1016 { $Tmp = "DeleteDiskVersion"}
					2001 { $Tmp = "RunAddDeviceToDomain"}
					2002 { $Tmp = "RunApplyAutoUpdate"}
					2003 { $Tmp = "RunApplyIncrementalUpdate"}
					2004 { $Tmp = "RunArchiveAuditTrail"}
					2005 { $Tmp = "RunAssignAuthGroup"}
					2006 { $Tmp = "RunAssignDevice"}
					2007 { $Tmp = "RunAssignDiskLocator"}
					2008 { $Tmp = "RunAssignServer"}
					2009 { $Tmp = "RunBoot"}
					2010 { $Tmp = "RunCopyPasteDevice"}
					2011 { $Tmp = "RunCopyPasteDisk"}
					2012 { $Tmp = "RunCopyPasteServer"}
					2013 { $Tmp = "RunCreateDirectory"}
					2014 { $Tmp = "RunCreateDiskCancel"}
					2015 { $Tmp = "RunDisableCollection"}
					2016 { $Tmp = "RunDisableDevice"}
					2017 { $Tmp = "RunDisableDeviceDiskLocator"}
					2018 { $Tmp = "RunDisableDiskLocator"}
					2019 { $Tmp = "RunDisableUserGroup"}
					2020 { $Tmp = "RunDisableUserGroupDiskLocator"}
					2021 { $Tmp = "RunDisplayMessage"}
					2022 { $Tmp = "RunEnableCollection"}
					2023 { $Tmp = "RunEnableDevice"}
					2024 { $Tmp = "RunEnableDeviceDiskLocator"}
					2025 { $Tmp = "RunEnableDiskLocator"}
					2026 { $Tmp = "RunEnableUserGroup"}
					2027 { $Tmp = "RunEnableUserGroupDiskLocator"}
					2028 { $Tmp = "RunExportOemLicenses"}
					2029 { $Tmp = "RunImportDatabase"}
					2030 { $Tmp = "RunImportDevices"}
					2031 { $Tmp = "RunImportOemLicenses"}
					2032 { $Tmp = "RunMarkDown"}
					2033 { $Tmp = "RunReboot"}
					2034 { $Tmp = "RunRemoveAuthGroup"}
					2035 { $Tmp = "RunRemoveDevice"}
					2036 { $Tmp = "RunRemoveDeviceFromDomain"}
					2037 { $Tmp = "RunRemoveDirectory"}
					2038 { $Tmp = "RunRemoveDiskLocator"}
					2039 { $Tmp = "RunResetDeviceForDomain"}
					2040 { $Tmp = "RunResetDatabaseConnection"}
					2041 { $Tmp = "RunRestartStreamingService"}
					2042 { $Tmp = "RunShutdown"}
					2043 { $Tmp = "RunStartStreamingService"}
					2044 { $Tmp = "RunStopStreamingService"}
					2045 { $Tmp = "RunUnlockAllDisk"}
					2046 { $Tmp = "RunUnlockDisk"}
					2047 { $Tmp = "RunServerStoreVolumeAccess"}
					2048 { $Tmp = "RunServerStoreVolumeMode"}
					2049 { $Tmp = "RunMergeDisk"}
					2050 { $Tmp = "RunRevertDiskVersion"}
					2051 { $Tmp = "RunPromoteDiskVersion"}
					2052 { $Tmp = "RunCancelDiskMaintenance"}
					2053 { $Tmp = "RunActivateDevice"}
					2054 { $Tmp = "RunAddDiskVersion"}
					2055 { $Tmp = "RunExportDisk"}
					2056 { $Tmp = "RunAssignDisk"}
					2057 { $Tmp = "RunRemoveDisk"}
					2057 { $Tmp = "RunDiskUpdateStart"}
					2057 { $Tmp = "RunDiskUpdateCancel"}
					2058 { $Tmp = "RunSetOverrideVersion"}
					2059 { $Tmp = "RunCancelTask"}
					2060 { $Tmp = "RunClearTask"}
					3001 { $Tmp = "RunWithReturnCreateDisk"}
					3002 { $Tmp = "RunWithReturnCreateDiskStatus"}
					3003 { $Tmp = "RunWithReturnMapDisk"}
					3004 { $Tmp = "RunWithReturnRebalanceDevices"}
					3005 { $Tmp = "RunWithReturnCreateMaintenanceVersion"}
					3006 { $Tmp = "RunWithReturnImportDisk"}
					4001 { $Tmp = "RunByteArrayInputImportDevices"}
					4002 { $Tmp = "RunByteArrayInputImportOemLicenses"}
					5001 { $Tmp = "RunByteArrayOutputArchiveAuditTrail"}
					5002 { $Tmp = "RunByteArrayOutputExportOemLicenses"}
					6001 { $Tmp = "SetAuthGroup"}
					6002 { $Tmp = "SetCollection"}
					6003 { $Tmp = "SetDevice"}
					6004 { $Tmp = "SetDisk"}
					6005 { $Tmp = "SetDiskLocator"}
					6006 { $Tmp = "SetFarm"}
					6007 { $Tmp = "SetFarmView"}
					6008 { $Tmp = "SetServer"}
					6009 { $Tmp = "SetServerBiosBootstrap"}
					6010 { $Tmp = "SetServerBootstrap"}
					6011 { $Tmp = "SetServerStore"}
					6012 { $Tmp = "SetSite"}
					6013 { $Tmp = "SetSiteView"}
					6014 { $Tmp = "SetStore"}
					6015 { $Tmp = "SetUserGroup"}
					6016 { $Tmp = "SetVirtualHostingPool"}
					6017 { $Tmp = "SetUpdateTask"}
					6018 { $Tmp = "SetDiskUpdateDevice"}
					7001 { $Tmp = "SetListDeviceBootstraps"}
					7002 { $Tmp = "SetListDeviceBootstrapsDelete"}
					7003 { $Tmp = "SetListDeviceBootstrapsAdd"}
					7004 { $Tmp = "SetListDeviceCustomProperty"}
					7005 { $Tmp = "SetListDeviceCustomPropertyDelete"}
					7006 { $Tmp = "SetListDeviceCustomPropertyAdd"}
					7007 { $Tmp = "SetListDeviceDiskPrinters"}
					7008 { $Tmp = "SetListDeviceDiskPrintersDelete"}
					7009 { $Tmp = "SetListDeviceDiskPrintersAdd"}
					7010 { $Tmp = "SetListDevicePersonality"}
					7011 { $Tmp = "SetListDevicePersonalityDelete"}
					7012 { $Tmp = "SetListDevicePersonalityAdd"}
					7013 { $Tmp = "SetListDevicePortBlockerCategories"}
					7014 { $Tmp = "SetListDevicePortBlockerCategoriesDelete"}
					7015 { $Tmp = "SetListDevicePortBlockerCategoriesAdd"}
					7016 { $Tmp = "SetListDevicePortBlockerOverrides"}
					7017 { $Tmp = "SetListDevicePortBlockerOverridesDelete"}
					7018 { $Tmp = "SetListDevicePortBlockerOverridesAdd"}
					7019 { $Tmp = "SetListDiskLocatorCustomProperty"}
					7020 { $Tmp = "SetListDiskLocatorCustomPropertyDelete"}
					7021 { $Tmp = "SetListDiskLocatorCustomPropertyAdd"}
					7022 { $Tmp = "SetListDiskLocatorPortBlockerCategories"}
					7023 { $Tmp = "SetListDiskLocatorPortBlockerCategoriesDelete"}
					7024 { $Tmp = "SetListDiskLocatorPortBlockerCategoriesAdd"}
					7025 { $Tmp = "SetListDiskLocatorPortBlockerOverrides"}
					7026 { $Tmp = "SetListDiskLocatorPortBlockerOverridesDelete"}
					7027 { $Tmp = "SetListDiskLocatorPortBlockerOverridesAdd"}
					7028 { $Tmp = "SetListServerCustomProperty"}
					7029 { $Tmp = "SetListServerCustomPropertyDelete"}
					7030 { $Tmp = "SetListServerCustomPropertyAdd"}
					7031 { $Tmp = "SetListUserGroupCustomProperty"}
					7032 { $Tmp = "SetListUserGroupCustomPropertyDelete"}
					7033 { $Tmp = "SetListUserGroupCustomPropertyAdd"}				
					Default {$Tmp = "Unknown"}
				}
				$TmpAction = $Tmp
				$Tmp = ""
				Switch ($Audit.type)
				{
					0 {$Tmp = "Many"}
					1 {$Tmp = "AuthGroup"}
					2 {$Tmp = "Collection"}
					3 {$Tmp = "Device"}
					4 {$Tmp = "Disk"}
					5 {$Tmp = "DeskLocator"}
					6 {$Tmp = "Farm"}
					7 {$Tmp = "FarmView"}
					8 {$Tmp = "Server"}
					9 {$Tmp = "Site"}
					10 {$Tmp = "SiteView"}
					11 {$Tmp = "Store"}
					12 {$Tmp = "System"}
					13 {$Tmp = "UserGroup"}
					Default { {$Tmp = "Undefined"}}
				}
				$TmpType = $Tmp
				$Tmp = $Null
				If($MSWord -or $PDF)
				{
					## Add the required key/values to the hashtable
					$WordTableRowHash = @{ DateTime=$Audit.time; Action=$TmpAction; Type=$TmpType; Name=$Audit.objectName; User=$Audit.userName; Path=$Audit.path;  }

					## Add the hash to the array
					$AuditWordTable += $WordTableRowHash;

					$CurrentServiceIndex++;
				}
				ElseIf($Text)
				{
				}
				ElseIf($HTML)
				{
				}
			}

			If($MSWord -or $PDF)
			{
				## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
				$Table = AddWordTable -Hashtable $AuditWordTable -Columns DateTime,Action,Type,Name,User,Path -Headers "Date/Time","Action","Type","Name","User","Path" -AutoFit $wdAutoFitContent;

				#set table to 9 point
				SetWordCellFormat -Collection $Table -Size 9
				## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				## IB - set column widths without recursion
				$Table.Columns.Item(1).Width = 65;
				$Table.Columns.Item(2).Width = 150;
				$Table.Columns.Item(3).Width = 55;
				$Table.Columns.Item(4).Width = 75;
				$Table.Columns.Item(5).Width = 75;
				$Table.Columns.Item(6).Width = 90;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$TableRange = $Null
				$Table = $Null
				Write-Verbose "$(Get-Date):"
			}
			ElseIf($Text)
			{
			}
			ElseIf($HTML)
			{
			}
		}
	}
}
Write-Verbose "$(Get-Date): "

$PVSSites            = $Null
$authgroups          = $Null
$servers             = $Null
$stores              = $Null
$bootstrapnames      = $Null
$tempserverbootstrap = $Null
$serverbootstraps    = $Null
$UserGroups          = $Null
$Disks               = $Null
$vDisks              = $Null
$Members             = $Null
$SiteViews           = $Null

#process the farm views now
Write-Verbose "$(Get-Date): Processing all PVS Farm Views"
$selection.InsertNewPage()
WriteWordLine 1 0 "Farm Views"
$Temp = $PVSSite.siteName
$GetWhat = "FarmView"
$GetParam = ""
$ErrorTxt = "Farm View information"
$FarmViews = BuildPVSObject $GetWhat $GetParam $ErrorTxt
If($FarmViews -ne $Null)
{
	ForEach($FarmView in $FarmViews)
	{
		Write-Verbose "$(Get-Date): `tProcessing Farm View $($FarmView.farmViewName)"
		Write-Verbose "$(Get-Date): `t`tProcessing General Tab"
		WriteWordLine 2 0 $FarmView.farmViewName
		WriteWordLine 0 1 "View Properties"
		WriteWordLine 0 2 "General"
		If(![String]::IsNullOrEmpty($FarmView.description))
		{
			WriteWordLine 0 3 "Name`t`t: " $FarmView.farmViewName
			WriteWordLine 0 3 "Description`t: " $FarmView.description
		}
		Else
		{
			WriteWordLine 0 3 "Name: " $FarmView.farmViewName
		}
		
		Write-Verbose "$(Get-Date): `t`tProcessing Members Tab"
		WriteWordLine 0 2 "Members"
		#process each target device contained in the farm view
		$Temp = $FarmView.farmViewId
		$GetWhat = "Device"
		$GetParam = "farmViewId = $Temp"
		$ErrorTxt = "Farm View Device Members information"
		$Members = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		If($Members -ne $Null)
		{
			ForEach($Member in $Members)
			{
				WriteWordLine 0 3 $Member.deviceName
			}
		}
	}
}
Else
{
	WriteWordLine 0 1 "There are no Farm Views configured"
}
Write-Verbose "$(Get-Date): "
$FarmViews = $Null
$Members = $Null

#process the stores now
Write-Verbose "$(Get-Date): Processing Stores"
$selection.InsertNewPage()
WriteWordLine 1 0 "Stores Properties"
$GetWhat = "Store"
$GetParam = ""
$ErrorTxt = "Farm Store information"
$Stores = BuildPVSObject $GetWhat $GetParam $ErrorTxt
If($Stores -ne $Null)
{
	ForEach($Store in $Stores)
	{
		Write-Verbose "$(Get-Date): `tProcessing Store $($Store.StoreName)"
		Write-Verbose "$(Get-Date): `t`tProcessing General Tab"
		WriteWordLine 2 0 $Store.StoreName
		WriteWordLine 0 1 "General"
		WriteWordLine 0 2 "Name`t`t: " $Store.StoreName
		If(![String]::IsNullOrEmpty($Store.description))
		{
			WriteWordLine 0 2 "Description`t: " $Store.description
		}
		
		WriteWordLine 0 2 "Store owner`t: " -nonewline
		If([String]::IsNullOrEmpty($Store.siteName))
		{
			WriteWordLine 0 0 "<none>"
		}
		Else
		{
			WriteWordLine 0 0 $Store.siteName
		}
		
		Write-Verbose "$(Get-Date): `t`tProcessing Servers Tab"
		WriteWordLine 0 1 "Servers"
		#find the servers (and the site) that serve this store
		$GetWhat = "Server"
		$GetParam = ""
		$ErrorTxt = "Server information"
		$Servers = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		$StoreSite = ""
		$StoreServers = @()
		If($Servers -ne $Null)
		{
			ForEach($Server in $Servers)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing Server $($Server.serverName)"
				$Temp = $Server.serverName
				$GetWhat = "ServerStore"
				$GetParam = "serverName = $Temp"
				$ErrorTxt = "Server Store information"
				$ServerStore = BuildPVSObject $GetWhat $GetParam $ErrorTxt
				If($ServerStore -ne $Null -and $ServerStore.storeName -eq $Store.StoreName)
				{
					$StoreSite = $Server.siteName
					$StoreServers +=  $Server.serverName
				}
			}	
		}
		WriteWordLine 0 2 "Site: " $StoreSite
		WriteWordLine 0 2 "Servers that provide this store:"
		ForEach($StoreServer in $StoreServers)
		{
			WriteWordLine 0 3 $StoreServer
		}

		Write-Verbose "$(Get-Date): `t`tProcessing Paths Tab"
		WriteWordLine 0 1 "Paths"
		WriteWordLine 0 2 "Default store path: " $Store.path
		If(![String]::IsNullOrEmpty($Store.cachePath))
		{
			WriteWordLine 0 2 "Default write-cache paths: "
			$WCPaths = $Store.cachePath.replace(",","`n`t`t`t")
			WriteWordLine 0 3 $WCPaths		
		}
	}
}
Else
{
	WriteWordLine 0 1 "There are no Stores configured"
}
Write-Verbose "$(Get-Date): "
$Stores = $Null
$Servers = $Null
$StoreSite = $Null
$StoreServers = $Null
$ServerStore = $Null

Write-Verbose "$(Get-Date): Create Appendix A Advanced Server Items (Server/Network)"
Write-Verbose "$(Get-Date): `t`tAdd Advanced Server Items table to doc"
If($MSWord -or $PDF)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Appendix A - Advanced Server Items (Server/Network)"
	## Create an array of hashtables to store our services
	[System.Collections.Hashtable[]] $ItemsWordTable = @();
	## Seed the row index from the second row
	[int] $CurrentServiceIndex = 2;
}
ElseIf($Text)
{
	Line 0 "Appendix A - Advanced Server Items (Server/Network)"
	Line 0 ""
}
ElseIf($HTML)
{
}

ForEach($Item in $AdvancedItems1)
{
	If($MSWord -or $PDF)
	{
		## Add the required key/values to the hashtable
		$WordTableRowHash = @{ 
		ServerName = $Item.serverName; 
		ThreadsperPort = $Item.threadsPerPort; 
		BuffersperThread = $Item.buffersPerThread; 
		ServerCacheTimeout = $Item.serverCacheTimeout; 
		LocalConcurrentIOLimit = $Item.localConcurrentIoLimit; 
		RemoteConcurrentIOLimit = $Item.remoteConcurrentIoLimit; 
		EthernetMTU = $Item.maxTransmissionUnits; 
		IOBurstSize = $Item.ioBurstSize; 
		EnableNonblockingIO = $Item.nonBlockingIoEnabled}

		## Add the hash to the array
		$ItemsWordTable += $WordTableRowHash;

		$CurrentServiceIndex++;
	}
	ElseIf($Text)
	{
		Line 1 "Server Name: " $Item.serverName
		Line 1 "Threads per Port: " $Item.threadsPerPort
		Line 1 "Buffers per Thread: " $Item.buffersPerThread
		Line 1 "Server Cache Timeout: " $Item.serverCacheTimeout
		Line 1 "Local Concurrent IO Limit: " $Item.localConcurrentIoLimit
		Line 1 "Remote Concurrent IO Limit: " $Item.remoteConcurrentIoLimit
		Line 1 "Ethernet MTU: " $Item.maxTransmissionUnits
		Line 1 "IO Burst Size: " $Item.ioBurstSize
		Line 1 "Enable Non-blocking IO: " $Item.nonBlockingIoEnabled
		Line 0 ""
	}
	ElseIf($HTML)
	{
	}
}

If($MSWord -or $PDF)
{
	## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	$Table = AddWordTable -Hashtable $ItemsWordTable `
	-Columns ServerName, ThreadsperPort, BuffersperThread, ServerCacheTimeout, LocalConcurrentIOLimit, RemoteConcurrentIOLimit, EthernetMTU, IOBurstSize, EnableNonblockingIO `
	-Headers "Server Name", "Threads per Port", "Buffers per Thread", "Server Cache Timeout", "Local Concurrent IO Limit", "Remote Concurrent IO Limit", "Ethernet MTU", "IO Burst Size", "Enable Non-blocking IO" `
	-AutoFit $wdAutoFitContent;

	## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
	SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	FindWordDocumentEnd
	$TableRange = $Null
	$Table = $Null
}
ElseIf($Text)
{
	#nothing to do
}
ElseIf($HTML)
{
}

Write-Verbose "$(Get-Date): `tFinished Creating Appendix A - Advanced Server Items (Server/Network)"
Write-Verbose "$(Get-Date): "

Write-Verbose "$(Get-Date): Create Appendix B Advanced Server Items (Pacing/Device)"
Write-Verbose "$(Get-Date): `t`tAdd Advanced Server Items table to doc"

If($MSWord -or $PDF)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Appendix B - Advanced Server Items (Pacing/Device)"
	## Create an array of hashtables to store our services
	[System.Collections.Hashtable[]] $ItemsWordTable = @();
	## Seed the row index from the second row
	[int] $CurrentServiceIndex = 2;
}
ElseIf($Text)
{
	Line 0 "Appendix B - Advanced Server Items (Pacing/Device)"
	Line 0 ""
}
ElseIf($HTML)
{
}

ForEach($Item in $AdvancedItems2)
{
	If($MSWord -or $PDF)
	{
		## Add the required key/values to the hashtable
		$WordTableRowHash = @{ 
		ServerName = $Item.serverName; 
		BootPauseSeconds = $Item.bootPauseSeconds; 
		MaximumBootTime = $Item.maxBootSeconds; 
		MaximumDevicesBooting = $Item.maxBootDevicesAllowed; 
		vDiskCreationPacing = $Item.vDiskCreatePacing; 
		LicenseTimeout = $Item.licenseTimeout}

		## Add the hash to the array
		$ItemsWordTable += $WordTableRowHash;

		$CurrentServiceIndex++;
	}
	ElseIf($Text)
	{
		Line 1 "Server Name: " $Item.serverName
		Line 1 "Boot Pause Seconds: " $Item.bootPauseSeconds
		Line 1 "Maximum Boot Time: " $Item.maxBootSeconds
		Line 1 "Maximum Devices Booting: " $Item.maxBootDevicesAllowed
		Line 1 "vDisk Creation Pacing: " $Item.vDiskCreatePacing
		Line 1 "License Timeout: " $Item.licenseTimeout
		Line 0 ""
	}
	ElseIf($HTML)
	{
	}
}

If($MSWord -or $PDF)
{
	## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	$Table = AddWordTable -Hashtable $ItemsWordTable `
	-Columns ServerName, BootPauseSeconds, MaximumBootTime, MaximumDevicesBooting, vDiskCreationPacing, LicenseTimeout `
	-Headers "Server Name", "Boot Pause Seconds", "Maximum Boot Time", "Maximum Devices Booting", "vDisk Creation Pacing", "License Timeout" `
	-AutoFit $wdAutoFitContent;

	## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
	SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	FindWordDocumentEnd
	$Table = $Null
}
ElseIf($Text)
{
	#nothing to do
}
ElseIf($HTML)
{
}

Write-Verbose "$(Get-Date): `tFinished Creating Appendix B - Advanced Server Items (Pacing/Device)"
Write-Verbose "$(Get-Date): "

Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

$AbstractTitle = "Citrix Provisioning Services Inventory"
$SubjectTitle = "Citrix Provisioning Services Inventory"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

If($MSWORD -or $PDF)
{
    SaveandCloseDocumentandShutdownWord
}
ElseIf($Text)
{
    SaveandCloseTextDocument
}
ElseIf($HTML)
{
    SaveandCloseHTMLDocument
}

Write-Verbose "$(Get-Date): Script has completed"
Write-Verbose "$(Get-Date): "

If($PDF)
{
	If(Test-Path "$($Script:FileName2)")
	{
		Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
	}
	Else
	{
		Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
		Write-Error "Unable to save the output file, $($Script:FileName2)"
	}
}
Else
{
	If(Test-Path "$($Script:FileName1)")
	{
		Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
	}
	Else
	{
		Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
		Write-Error "Unable to save the output file, $($Script:FileName1)"
	}
}

Write-Verbose "$(Get-Date): "

#http://poshtips.com/measuring-elapsed-time-in-powershell/
Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
$runtime = $(Get-Date) - $Script:StartTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	$runtime.Days, `
	$runtime.Hours, `
	$runtime.Minutes, `
	$runtime.Seconds,
	$runtime.Milliseconds)
Write-Verbose "$(Get-Date): Elapsed time: $($Str)"
$runtime = $Null
$Str = $Null
$ErrorActionPreference = $SaveEAPreference
