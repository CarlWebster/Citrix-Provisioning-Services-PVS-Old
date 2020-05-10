<#
.SYNOPSIS
	Creates a complete inventory of a Citrix PVS 5.x, 6.x farm using Microsoft Word.
.DESCRIPTION
	Creates a complete inventory of a Citrix PVS 5.x, 6.x farm using Microsoft Word and PowerShell.
	Creates a Word document named after the PVS 5.x, 6.x farm.
	Document includes a Cover Page, Table of Contents and Footer.
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\Company
	This parameter has an alias of CN.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	(default cover pages in Word en-US)
	Valid input is:
		Alphabet (Word 2007/2010. Works)
		Annual (Word 2007/2010. Doesn't really work well for this report)
		Austere (Word 2007/2010. Works)
		Austin (Word 2010/2013. Doesn't work in 2013, mostly works in 2007/2010 but Subtitle/Subject & Author fields need to me moved after title box is moved up)
		Banded (Word 2013. Works)
		Conservative (Word 2007/2010. Works)
		Contrast (Word 2007/2010. Works)
		Cubicles (Word 2007/2010. Works)
		Exposure (Word 2007/2010. Works if you like looking sideways)
		Facet (Word 2013. Works)
		Filigree (Word 2013. Works)
		Grid (Word 2010/2013.Works in 2010)
		Integral (Word 2013. Works)
		Ion (Dark) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Ion (Light) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Mod (Word 2007/2010. Works)
		Motion (Word 2007/2010/2013. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2007/2010. Works)
		Puzzle (Word 2007/2010. Top date doesn't fit, box needs to be manually resized or font changed to 14 point)
		Retrospect (Word 2013. Works)
		Semaphore (Word 2013. Works)
		Sideline (Word 2007/2010/2013. Doesn't work in 2013, works in 2007/2010)
		Slice (Dark) (Word 2013. Doesn't work)
		Slice (Light) (Word 2013. Doesn't work)
		Stacks (Word 2007/2010. Works)
		Tiles (Word 2007/2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2007/2010. Works)
		ViewMaster (Word 2013. Works)
		Whisp (Word 2013. Works)
	Default value is Motion.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER AdminAddress
	Specifies the name of a PVS server that the PowerShell script will connect to. 
.PARAMETER User
	Specifies the user used for the AdminAddress connection. 
.PARAMETER Domain
	Specifies the domain used for the AdminAddress connection. 
.PARAMETER Password
	Specifies the password used for the AdminAddress connection. 
.EXAMPLE
	PS C:\PSScript > .\PVS_Inventory_v2.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator
	AdminAddress = LocalHost

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
	LocalHost for AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\PVS_Inventory_v2.ps1 -verbose
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator
	AdminAddress = LocalHost

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
	LocalHost for AdminAddress.
	Will display verbose messages as the script is running.
.EXAMPLE
	PS C:\PSScript .\PVS_Inventory_v2.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\PVS_Inventory_v2.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster" -AdminAddress PVS1

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
		PVS1 for AdminAddress.
.EXAMPLE
	PS C:\PSScript .\PVS_Inventory_v2.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster" -AdminAddress PVS1 -User cwebster -Domain WebstersLab -Password Abc123!@#

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
		PVS1 for AdminAddress.
		cwebster for User.
		WebstersLab for Domain.
		Abc123!@# for Password.
.EXAMPLE
	PS C:\PSScript .\PVS_Inventory_v2.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster" -AdminAddress PVS1 -User cwebster

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
		PVS1 for AdminAddress.
		cwebster for User.
		Script will prompt for the Domain and Password
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word document.
.LINK
	http://www.carlwebster.com/documenting-a-citrix-provisioning-services-farm-with-microsoft-powershell-and-word-version-2
.NOTES
	NAME: PVS_Inventory_V2.ps1
	VERSION: 2.03
	AUTHOR: Carl Webster (with a lot of help from Michael B. Smith and Jeff Wouters)
	LASTEDIT: June 17, 2013
#>


#thanks to @jeffwouters for helping me with these parameters
[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "" ) ]

Param([parameter(
	Position = 0, 
	Mandatory=$false )
	] 
	[Alias("CN")]
	[string]$CompanyName="",
    
	[parameter(
	Position = 1, 
	Mandatory=$false )
	] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Motion", 

	[parameter(
	Position = 2, 
	Mandatory=$false )
	] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,
		
	[parameter(
	Position = 3, 
	Mandatory=$false )
	] 
	[string]$AdminAddress="",

	[parameter(
	Position = 4, 
	Mandatory=$false )
	] 
	[string]$User="",

	[parameter(
	Position = 5, 
	Mandatory=$false )
	] 
	[string]$Domain="",

	[parameter(
	Position = 6, 
	Mandatory=$false )
	] 
	[string]$Password="")

Set-StrictMode -Version 2

#Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#This script written for "Benji", March 19, 2012
#Thanks to Michael B. Smith, Joe Shonk and Stephane Thirion
#for testing and fine-tuning tips 
#Updated Janury 29, 2013 to create a Word 2007/2010/2013 document
#	Fixed numerous bugs and logic issues
#	Test for CompanyName in two different registry locations
#	Fixed issues found by running in set-strictmode -version 2.0
#	Fixed typos
#	Test if template DOTX file loads properly.  If not, skip Cover Page and Table of Contents
#	Add more write-verbose statements
#	Disable Spell and Grammer Check to resolve issue and improve performance (from Pat Coughlin)
#Updated March 14, 2013
#	?{?_.SessionId -eq $SessionID} should have been ?{$_.SessionId -eq $SessionID} in the CheckWordPrereq function
#Updated March 16, 2013
#	Fixed hard coded "6.5" in report subject.  Copy and Paste error from the XenApp 6.5 script.
#Updated April 19, 2013
#	Fixed the content of and the detail contained in the Table of Contents
#	Fixed a compatibility issue with the way the Word file was saved and Set-StrictMode -Version 2
#Updated June 7, 2013
#	Added for PVS 6.x processing the vDisk Load Balancing menu (bug found by Corey Tracey)
#Updated June 17, 2013
#	Added three command line parameters for use with -AdminAddress (User, Domain, Password) at the request of Corey Tracey


Function CheckWordPrereq
{
	if ((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		Write-Host "This script directly outputs to Microsoft Word, please install Microsoft Word"
		exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $null
	if ($wordrunning)
	{
		Write-Host "Please close all instances of Microsoft Word before running this report."
		exit
	}
}

Function BuildPVSObject
{
	Param( [string]$MCLIGetWhat = '', [string]$MCLIGetParameters = '', [string]$TextForErrorMsg = '' )

	$error.Clear()

	If($MCLIGetParameters -ne '')
	{
		$MCLIGetResult = Mcli-Get "$($MCLIGetWhat)" -p "$($MCLIGetParameters)"
	}
	Else
	{
		$MCLIGetResult = Mcli-Get "$($MCLIGetWhat)"
	}

	If( $error.Count -eq 0 )
	{
		$PluralObject = @()
		$SingleObject = $null
		foreach( $record in $MCLIGetResult )
		{
			If($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
			{
				If($SingleObject -ne $null)
				{
					$PluralObject += $SingleObject
				}
				$SingleObject = new-object System.Object
			}

			$index = $record.IndexOf( ':' )
			if( $index -gt 0 )
			{
				$property = $record.SubString( 0, $index  )
				$value    = $record.SubString( $index + 2 )
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
	Param( $xDevice )

	If($xDevice -eq $null -or $xDevice.status -eq "" -or $xDevice.status -eq "0")
	{
		WriteWordLine 0 3 "Target device inactive"
	}
	Else
	{
		WriteWordLine 0 2 "Target device active"
		WriteWordLine 0 2 "IP Address: " $xDevice.ip
		WriteWordLine 0 2 "Server: " -nonewline
		WriteWordLine 0 0 "$($xDevice.serverName) `($($xDevice.serverIpConnection)`: $($xDevice.serverPortConnection)`)"
		WriteWordLine 0 2 "Retries: " $xDevice.status
		WriteWordLine 0 2 "vDisk: " $xDevice.diskLocatorName
		WriteWordLine 0 2 "vDisk version: " $xDevice.diskVersion
		WriteWordLine 0 2 "vDisk full name: " $xDevice.diskFileName
		WriteWordLine 0 2 "vDisk access: " -nonewline
		switch ($xDevice.diskVersionAccess)
		{
			0 {WriteWordLine 0 0 "Production"}
			1 {WriteWordLine 0 0 "Test"}
			2 {WriteWordLine 0 0 "Maintenance"}
			3 {WriteWordLine 0 0 "Personal vDisk"}
			Default {WriteWordLine 0 0 "vDisk access type could not be determined: $($xDevice.diskVersionAccess)"}
		}
		switch($xDevice.licenseType)
		{
			0 {WriteWordLine 0 2 "No License"}
			1 {WriteWordLine 0 2 "Desktop License"}
			2 {WriteWordLine 0 2 "Server License"}
			5 {WriteWordLine 0 2 "OEM SmartClient License"}
			6 {WriteWordLine 0 2 "XenApp License"}
			7 {WriteWordLine 0 2 "XenDesktop License"}
			Default {WriteWordLine 0 2 "Device license type could not be determined: $($xDevice.licenseType)"}
		}
		
		WriteWordLine 0 1 "Logging"
		WriteWordLine 0 2 "Logging level: " -nonewline
		switch ($xDevice.logLevel)
		{
			0   {WriteWordLine 0 0 "Off"    }
			1   {WriteWordLine 0 0 "Fatal"  }
			2   {WriteWordLine 0 0 "Error"  }
			3   {WriteWordLine 0 0 "Warning"}
			4   {WriteWordLine 0 0 "Info"   }
			5   {WriteWordLine 0 0 "Debug"  }
			6   {WriteWordLine 0 0 "Trace"  }
			default {WriteWordLine 0 0 "Logging level could not be determined: $($xDevice.logLevel)"}
		}
		WriteWordLine 0 0 ""
	}
}

Function ValidateCompanyName
{
	$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
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
# This function just gets $true or $false
function Test-RegistryValue($path, $name)
{
    $key = Get-Item -LiteralPath $path -EA 0
    $key -and $null -ne $key.GetValue($name, $null)
}

# Gets the specified registry value or $null if it is missing
function Get-RegistryValue($path, $name)
{
    $key = Get-Item -LiteralPath $path -EA 0
    if ($key) {
        $key.GetValue($name, $null)
    }
}

Function ValidateCoverPage
{
	Param( [int]$xWordVersion, [string]$xCP )
	
	$xArray = ""
	If( $xWordVersion -eq 15)
	{
		#word 2013
		$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect", "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
	}
	ElseIf( $xWordVersion -eq 14)
	{
		#word 2010
		$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative", "Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint", "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
	}
	ElseIf( $xWordVersion -eq 12)
	{
		#word 2007
		$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast", "Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend" )
	}
	
	If ($xArray -contains $xCP)
	{
		Return $True
	}
	Else
	{
		Return $False
	}
}

Function Check-NeededPSSnapins
{
	Param( [parameter(Mandatory = $true)][alias("Snapin")][string[]]$Snapins)
	
    #function specifics
    $MissingSnapins=@()
    $FoundMissingSnapin=$false
    $loadedSnapins = @()
    $registeredSnapins = @()
    
    #Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
    $loadedSnapins += get-pssnapin | % {$_.name}
    $registeredSnapins += get-pssnapin -Registered | % {$_.name}
    
    
    foreach ($Snapin in $Snapins){
        #check if the snapin is loaded
        if (!($LoadedSnapins -like $snapin)){

            #Check if the snapin is missing
            if (!($RegisteredSnapins -like $Snapin)){

                #set the flag if it's not already
                if (!($FoundMissingSnapin)){
                    $FoundMissingSnapin = $True
                }
                
                #add the entry to the list
                $MissingSnapins += $Snapin
            }#End Registered If 
            
            Else{
                #Snapin is registered, but not loaded, loading it now:
                Write-Host "Loading Windows PowerShell snap-in: $snapin"
                Add-PSSnapin -Name $snapin
            }
            
        }#End Loaded If
        #Snapin is registered and loaded
        else{write-debug "Windows PowerShell snap-in: $snapin - Already Loaded"}
    }#End For
    
    if ($FoundMissingSnapin){
        write-warning "Missing Windows PowerShell snap-ins Detected:"
        $missingSnapins | % {write-warning "($_)"}
        return $False
    }#End If
    
    Else{
        Return $true
    }#End Else
    
}#End Function

Function WriteWordLine
#function created by Ryan Revord
#@rsrevord on Twitter
#function created to make output to Word easy in this script
{
	Param( [int]$style=0, [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "'n", [switch]$nonewline)
	$output=""
	#Build output style
	switch ($style)
	{
		0 {$Selection.Style = "No Spacing"}
		1 {$Selection.Style = "Heading 1"}
		2 {$Selection.Style = "Heading 2"}
		3 {$Selection.Style = "Heading 3"}
		4 {$Selection.Style = "Heading 4"}
		5 {$Selection.Style = "Heading 5"}
		Default {$Selection.Style = "No Spacing"}
	}
	#build # of tabs
	While( $tabs -gt 0 ) { 
		$output += "`t"; $tabs--; 
	}
		
	#output the rest of the parameters.
	$output += $name + $value
	$Selection.TypeText($output)
	
	#test for new WriteWordLine
	If($nonewline){
		# Do nothing.
	} Else {
		$Selection.TypeParagraph()
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop=$properties | foreach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$null,$_,$null)
		if ($propname -eq $Name) 
		{
			Return $_
		}
	} #foreach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$null,$prop,$Value)
}

#script begins
write-verbose "checking for McliPSSnapin"
if (!(Check-NeededPSSnapins "McliPSSnapIn")){
    #We're missing Citrix Snapins that we need
    write-error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Script will now close."
    break
}

CheckWordPreReq

#setup remoting if $AdminAddress is not empty
If(![System.String]::IsNullOrEmpty( $AdminAddress ))
{
	If(![System.String]::IsNullOrEmpty( $User ))
	{
		If([System.String]::IsNullOrEmpty( $Domain ))
		{
			$Domain = Read-Host "Domain name for user is required.  Enter Domain name for user"
		}		

		If([System.String]::IsNullOrEmpty( $Password ))
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

	If( $error.Count -eq 0 )
	{
		Write-Verbose "This script is being run remotely against server $($AdminAddress)"
		If(![System.String]::IsNullOrEmpty( $User ))
		{
			Write-Verbose "User=$($User)"
			Write-Verbose "Domain=$($Domain)"
		}
	}
	Else 
	{
		Write-Warning "Remoting could not be setup to server $($AdminAddress)"
		Write-Warning "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
		Write-Warning "Script cannot continue"
		Exit
	}
}

#get PVS major version
write-verbose "Getting PVS version info"

$error.Clear()
$tempversion = mcli-info version
If( $error.Count -eq 0 )
{
	#build PVS version values
	$version = new-object System.Object 
	foreach( $record in $tempversion )
	{
		$index = $record.IndexOf( ':' )
		if( $index -gt 0 )
		{
			$property = $record.SubString( 0, $index)
			$value = $record.SubString( $index + 2 )
			Add-Member -inputObject $version -MemberType NoteProperty -Name $property -Value $value
		}
	}
} 
Else 
{
	Write-Warning "PVS version information could not be retrieved"
	Write-Warning "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
	Write-error "Script is terminating"
	#without version info, script should not proceed
	Exit
}

$PVSVersion     = $Version.mapiVersion.SubString(0,1)
$PVSFullVersion = $Version.mapiVersion.SubString(0,3)
$tempversion    = $null
$version        = $null

$FarmAutoAddEnabled = $false

#build PVS farm values
write-verbose "Build PVS farm values"
#there can only be one farm
$GetWhat = "Farm"
$GetParam = ""
$ErrorTxt = "PVS Farm information"
$farm = BuildPVSObject $GetWhat $GetParam $ErrorTxt

If($Farm -eq $null)
{
	#without farm info, script should not proceed
	write-error "PVS Farm information could not be retrieved.  Script is terminating."
	Break
}

$FarmName = $farm.FarmName
$Title="Inventory Report for the $($FarmName) Farm"
$filename="$($pwd.path)\$($farm.FarmName).docx"

write-verbose "Setting up Word"
#these values were attained from 
#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
$wdSeekPrimaryFooter = 4
$wdAlignPageNumberRight = 2
$wdStory = 6
$wdMove = 0
$wdSeekMainDocument = 0
$wdColorGray15 = 14277081

# Setup word for output
write-verbose "Create Word comObject.  If you are not running Word 2007, ignore the next message."
$Word = New-Object -comobject "Word.Application"
$WordVersion = [int] $Word.Version
If( $WordVersion -eq 15)
{
	write-verbose "Running Microsoft Word 2013"
	$WordProduct = "Word 2013"
}
Elseif ( $WordVersion -eq 14)
{
	write-verbose "Running Microsoft Word 2010"
	$WordProduct = "Word 2010"
}
Elseif ( $WordVersion -eq 12)
{
	write-verbose "Running Microsoft Word 2007"
	$WordProduct = "Word 2007"
}
Elseif ( $WordVersion -eq 11)
{
	write-verbose "Running Microsoft Word 2003"
	Write-error "This script does not work with Word 2003. Script will end."
	$word.quit()
	exit
}
Else
{
	Write-error "You are running an untested or unsupported version of Microsoft Word.  Script will end."
	$word.quit()
	exit
}

write-verbose "Validate company name"
#only validate CompanyName if the field is blank
If([String]::IsNullOrEmpty($CompanyName))
{
	$CompanyName = ValidateCompanyName
	If([String]::IsNullOrEmpty($CompanyName))
	{
		write-error "Company Name cannot be blank.  Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value.  Script cannot continue."
		$Word.Quit()
		exit
	}
}

write-verbose "Validate cover page"
$ValidCP = ValidateCoverPage $WordVersion $CoverPage
If(!$ValidCP)
{
	write-error "For $WordProduct, $CoverPage is not a valid Cover Page option.  Script cannot continue."
	$Word.Quit()
	exit
}

Write-Verbose "Company Name: $CompanyName"
Write-Verbose "Cover Page  : $CoverPage"
Write-Verbose "User Name   : $UserName"
Write-Verbose "Farm Name   : $FarmName"
Write-Verbose "Title       : $Title"
Write-Verbose "Filename    : $filename"

$Word.Visible = $False

#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
#using Jeff's Demo-WordReport.ps1 file for examples
#down to $global:configlog = $false is from Jeff Hicks
write-verbose "Load Word Templates"
$CoverPagesExist = $False
$word.Templates.LoadBuildingBlocks()
If ( $WordVersion -eq 12)
{
	#word 2007
	$BuildingBlocks=$word.Templates | Where {$_.name -eq "Building Blocks.dotx"}
}
Else
{
	#word 2010/2013
	$BuildingBlocks=$word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}
}

If($BuildingBlocks -ne $Null)
{
	$CoverPagesExist = $True
	$part=$BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
}
Else
{
	$CoverPagesExist = $False
}

write-verbose "Create empty word doc"
$Doc = $Word.Documents.Add()
$global:Selection = $Word.Selection

#Disable Spell and Grammer Check to resolve issue and improve performance (from Pat Coughlin)
write-verbose "disable spell checking"
$Word.Options.CheckGrammarAsYouType=$false
$Word.Options.CheckSpellingAsYouType=$false

If($CoverPagesExist)
{
	#insert new page, getting ready for table of contents
	write-verbose "insert new page, getting ready for table of contents"
	$part.Insert($selection.Range,$True) | out-null
	$selection.InsertNewPage()

	#table of contents
	write-verbose "table of contents"
	$toc=$BuildingBlocks.BuildingBlockEntries.Item("Automatic Table 2")
	$toc.insert($selection.Range,$True) | out-null
}
Else
{
	write-verbose "Cover Pages are not installed."
	write-warning "Cover Pages are not installed so this report will not have a cover page."
	write-verbose "Table of Contents are not installed."
	write-warning "Table of Contents are not installed so this report will not have a Table of Contents."
}

#set the footer
write-verbose "set the footer"
[string]$footertext="Report created by $username"

#get the footer
write-verbose "get the footer and format font"
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekPrimaryFooter
#get the footer and format font
$footers=$doc.Sections.Last.Footers
foreach ($footer in $footers) 
{
	if ($footer.exists) 
	{
		$footer.range.Font.name="Calibri"
		$footer.range.Font.size=8
		$footer.range.Font.Italic=$True
		$footer.range.Font.Bold=$True
	}
} #end Foreach
write-verbose "Footer text"
$selection.HeaderFooter.Range.Text=$footerText

#add page numbering
write-verbose "add page numbering"
$selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

#return focus to main document
write-verbose "return focus to main document"
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
write-verbose "move to the end of the current document"
$selection.EndKey($wdStory,$wdMove) | Out-Null
#end of Jeff Hicks 

write-verbose "Processing PVS Farm Information"
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
write-verbose "Processing Security Tab"
WriteWordLine 2 0 "Security"
WriteWordLine 0 1 "Groups with Farm Administrator access:"
#build security tab values
$GetWhat = "authgroup"
$GetParam = "farm=1"
$ErrorTxt = "Groups with Farm Administrator access"
$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

If($AuthGroups -ne $null)
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
write-verbose "Processing Groups Tab"
WriteWordLine 2 0 "Groups"
WriteWordLine 0 1 "All the Security Groups that can be assigned access rights:"
$GetWhat = "authgroup"
$GetParam = ""
$ErrorTxt = "Security Groups information"
$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

If($AuthGroups -ne $null)
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
write-verbose "Processing Licensing Tab"
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
write-verbose "Processing Options Tab"
WriteWordLine 2 0 "Options"
WriteWordLine 0 1 "Auto-Add"
WriteWordLine 0 2 "Enable auto-add: " -nonewline
If($farm.autoAddEnabled -eq "1")
{
	WriteWordLine 0 0 "Yes"
	WriteWordLine 0 3 "Add new devices to this site: " $farm.defaultSiteName
	$FarmAutoAddEnabled = $true
}
Else
{
	WriteWordLine 0 0 "No"	
	$FarmAutoAddEnabled = $false
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

If($PVSVersion -eq "6")
{
	#vDisk Version tab
	write-verbose "Processing vDisk Version Tab"
	WriteWordLine 2 0 "vDisk Version"
	WriteWordLine 0 1 "Alert if number of version from base image exceeds`t`t: " $farm.maxVersions
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
	switch ($Farm.mergeMode)
	{
		0   {WriteWordLine 0 0 "Production" }
		1   {WriteWordLine 0 0 "Test"       }
		2   {WriteWordLine 0 0 "Maintenance"}
		default {WriteWordLine 0 0 "Default access mode could not be determined: $($Farm.mergeMode)"}
	}
}

#status tab
write-verbose "Processing Status Tab"
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
	
$farm = $null
$authgroups = $null

#build site values
write-verbose "Processing Sites"
$GetWhat = "site"
$GetParam = ""
$ErrorTxt = "PVS Site information"
$PVSSites = BuildPVSObject $GetWhat $GetParam $ErrorTxt

ForEach($PVSSite in $PVSSites)
{
	write-verbose "Processing Site $($PVSSite.siteName)"
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
	write-verbose "Processing Security Tab"
	$temp = $PVSSite.SiteName
	$GetWhat = "authgroup"
	$GetParam = "sitename=$temp"
	$ErrorTxt = "Groups with Site Administrator access"
	$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	WriteWordLine 2 0 "Security"
	If($authGroups -ne $null)
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
	write-verbose "Processing Options Tab"
	WriteWordLine 2 0 "Options"
	WriteWordLine 0 1 "Auto-Add"
	If($PVSVersion -eq "5" -or ($PVSVersion -eq "6" -and $FarmAutoAddEnabled))
	{
		WriteWordLine 0 2 "Add new devices to this collection: " -nonewline
		If($PVSSite.defaultCollectionName)
		{
			WriteWordLine 0 0 $PVSSite.defaultCollectionName
		}
		Else
		{
			WriteWordLine 0 0 "<No default collection>"
		}
	}
	If($PVSVersion -eq "6")
	{
		WriteWordLine 0 2 "Seconds between vDisk inventory scans: " $PVSSite.inventoryFilePollingInterval

		#vDisk Update
		write-verbose "Processing vDisk Update Tab"
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
	write-verbose "Processing Servers in Site $($PVSSite.siteName)"
	$temp = $PVSSite.SiteName
	$GetWhat = "server"
	$GetParam = "sitename=$temp"
	$ErrorTxt = "Servers for Site $temp"
	$servers = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	WriteWordLine 2 0 "Servers"
	ForEach($Server in $Servers)
	{
		write-verbose "Processing Server $($Server.serverName)"
		#general tab
		write-verbose "Processing General Tab"
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
		IF($Server.eventLoggingEnabled -eq "1")
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
			
		write-verbose "Processing Network Tab"
		WriteWordLine 0 1 "Network"
		WriteWordLine 0 2 "IP addresses:"
		$test = $Server.ip.ToString()
		$test1 = $test.replace(",","`n`t`t`t")
		WriteWordLine 0 3 $test1
		WriteWordLine 0 2 "Ports"
		WriteWordLine 0 3 "First port: " $Server.firstPort
		WriteWordLine 0 3 "Last port: " $Server.lastPort
			
		write-verbose "Processing Stores Tab"
		WriteWordLine 0 1 "Stores"
		#process all stores for this server
		write-verbose "Processing Stores for server"
		$temp = $Server.serverName
		$GetWhat = "serverstore"
		$GetParam = "servername=$temp"
		$ErrorTxt = "Store information for server $temp"
		$stores = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		WriteWordLine 0 2 "Stores that this server supports:"

		If($Stores -ne $null)
		{
			ForEach($store in $stores)
			{
				write-verbose "Processing Store $($store.storename)"
				WriteWordLine 0 3 "Store`t: " $store.storename
				WriteWordLine 0 3 "Path`t: " -nonewline
				If($store.path.length -gt 0)
				{
					WriteWordLine 0 0 $store.path
				}
				Else
				{
					WriteWordLine 0 0 "<Using the default path from the store>"
				}
				WriteWordLine 0 3 "Write cache paths: " -nonewline
				If($store.cachePath.length -gt 0)
				{
					WriteWordLine 0 0 $store.cachePath
				}
				Else
				{
					WriteWordLine 0 0 "<Using the default path from the store>"
				}
				WriteWordLine 0 0 ""
			}
		}

		write-verbose "Processing Options Tab"
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
			
		write-verbose "Processing Logging Tab"
		WriteWordLine 0 1 "Logging"
		WriteWordLine 0 2 "Logging level: " -nonewline
		switch ($Server.logLevel)
		{
			0   {WriteWordLine 0 0 "Off"    }
			1   {WriteWordLine 0 0 "Fatal"  }
			2   {WriteWordLine 0 0 "Error"  }
			3   {WriteWordLine 0 0 "Warning"}
			4   {WriteWordLine 0 0 "Info"   }
			5   {WriteWordLine 0 0 "Debug"  }
			6   {WriteWordLine 0 0 "Trace"  }
			default {WriteWordLine 0 0 "Logging level could not be determined: $($Server.logLevel)"}
		}
		WriteWordLine 0 3 "File size maximum`t: $($Server.logFileSizeMax) (MB)"
		WriteWordLine 0 3 "Backup files maximum`t: " $Server.logFileBackupCopiesMax
		WriteWordLine 0 0 ""
		
		#advanced button at the bottom
		write-verbose "Processing Server Tab on Advanced button"
		WriteWordLine 0 1 "Advanced button"
		WriteWordLine 0 2 "Server"
		WriteWordLine 0 3 "Threads per port`t`t: " $Server.threadsPerPort
		WriteWordLine 0 3 "Buffers per thread`t`t: " $Server.buffersPerThread
		WriteWordLine 0 3 "Server cache timeout`t`t: $($Server.serverCacheTimeout) (seconds)"
		WriteWordLine 0 3 "Local concurrent I/O limit`t: $($Server.localConcurrentIoLimit) (transactions)"
		WriteWordLine 0 3 "Remote concurrent I/O limit`t: $($Server.remoteConcurrentIoLimit) (transactions)"

		write-verbose "Processing Network Tab on Advanced button"
		WriteWordLine 0 2 "Network"
		WriteWordLine 0 3 "Ethernet MTU`t: $($Server.maxTransmissionUnits) (bytes)"
		WriteWordLine 0 3 "I/O burst size`t: $($Server.ioBurstSize) (KB)"
		WriteWordLine 0 3 "Enable non-blocking I/O for network communications: " -nonewline
		If($Server.nonBlockingIoEnabled -eq "1")
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}

		write-verbose "Processing Pacing Tab on Advanced button"
		WriteWordLine 0 2 "Pacing"
		WriteWordLine 0 3 "Boot pause seconds`t`t: " $Server.bootPauseSeconds
		WriteWordLine 0 3 "Maximum boot time`t`t: $($Server.maxBootSeconds) (seconds)"
		WriteWordLine 0 3 "Maximum devices booting`t: " $Server.maxBootDevicesAllowed
		WriteWordLine 0 3 "vDisk Creation pacing`t`t: " $Server.vDiskCreatePacing

		write-verbose "Processing Device Tab on Advanced button"
		WriteWordLine 0 2 "Device"
		WriteWordLine 0 3 "License timeout: $($Server.licenseTimeout) (seconds)"

		WriteWordLine 0 0 ""
	}

	#the properties for the servers have been processed. 
	#now to process the stuff available via a right-click on each server

	#Configure Bootstrap is first
	write-verbose "Processing Bootstrap files"
	WriteWordLine 2 0 "Configure Bootstrap settings"
	ForEach($Server in $Servers)
	{
		write-verbose "Processing Bootstrap files for Server $($server.servername)"
		#first get all bootstrap files for the server
		$temp = $server.serverName
		$GetWhat = "ServerBootstrapNames"
		$GetParam = "serverName=$temp"
		$ErrorTxt = "Server Bootstrap Name information"
		$BootstrapNames = BuildPVSObject $GetWhat $GetParam $ErrorTxt

		#Now that the list of bootstrap names has been gathered
		#We have the mandatory parameter to get the bootstrap info
		#there should be at least one bootstrap filename
		WriteWordLine 3 0 $Server.serverName
		If($Bootstrapnames -ne $null)
		{
			#cannot use the BuildPVSObject function here
			$serverbootstraps=@()
			ForEach($Bootstrapname in $Bootstrapnames)
			{
				#get serverbootstrap info
				$error.Clear()
				$tempserverbootstrap = Mcli-Get ServerBootstrap -p name="$($Bootstrapname.name)",servername="$($server.serverName)"
				If( $error.Count -eq 0 )
				{
					$serverbootstrap = $null
					foreach( $record in $tempserverbootstrap )
					{
						If($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
						{
							If($serverbootstrap -ne $null)
							{
								$serverbootstraps += $serverbootstrap
							}
							$serverbootstrap = new-object System.Object
							#add the bootstrapname name value to the serverbootstrap object
							$property = "BootstrapName"
							$value = $Bootstrapname.name
							Add-Member -inputObject $serverbootstrap -MemberType NoteProperty -Name $property -Value $value
						}
						$index = $record.IndexOf( ':' )
						if( $index -gt 0 )
						{
							$property = $record.SubString( 0, $index)
							$value = $record.SubString( $index + 2 )
							If($property -ne "Executing")
							{
								Add-Member -inputObject $serverbootstrap -MemberType NoteProperty -Name $property -Value $value
							}
						}
					}
					$serverbootstraps += $serverbootstrap
				}
				Else
				{
					WriteWordLine 0 0 "Server Bootstrap information could not be retrieved"
					WriteWordLine 0 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
				}
			}
			If($ServerBootstraps -ne $null)
			{
				write-verbose "Processing General Tab"
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
			write-verbose "Processing Options Tab"
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

	#process all vDisks in site
	write-verbose "Processing all vDisks in site"
	$Temp = $PVSSite.SiteName
	$GetWhat = "DiskInfo"
	$GetParam = "siteName=$Temp"
	$ErrorTxt = "Disk information"
	$Disks = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	WriteWordLine 2 0 "vDisk Pool"
	If($Disks -ne $null)
	{
		ForEach($Disk in $Disks)
		{
			write-verbose "Processing vDisk $($Disk.diskLocatorName)"
			write-verbose "Processing Properties General Tab"
			WriteWordLine 3 0 $Disk.diskLocatorName
			If($PVSVersion -eq "5")
			{
				#PVS 5.x
				WriteWordLine 0 1 "vDisk Properties"
				WriteWordLine 0 2 "General"
				WriteWordLine 0 3 "Store`t`t`t: " $Disk.storeName
				WriteWordLine 0 3 "Site`t`t`t: " $Disk.siteName
				WriteWordLine 0 3 "Filename`t`t: " $Disk.diskLocatorName
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
					switch ($Disk.subnetAffinity)
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

				WriteWordLine 0 2 "Mode"
				WriteWordLine 0 3 "Access mode: " -nonewline
				If($Disk.writeCacheType -eq "0")
				{
					WriteWordLine 0 0 "Private Image (single device, read/write access)"
				}
				Elseif ($Disk.writeCacheType -eq "7")
				{
					WriteWordLine 0 0 "Difference Disk Image"
				}
				Else
				{
					WriteWordLine 0 0 "Standard Image (multi-device, read-only access)"
					WriteWordLine 0 3 "Cache type: " -nonewline
					If($PVSVersion -eq "6")
					{
						switch ($Disk.writeCacheType)
						{
							0   {WriteWordLine 0 0 "Private Image"}
							1   {WriteWordLine 0 0 "Cache on server"}
							3   {WriteWordLine 0 0 "Cache in device RAM"}
							4   {WriteWordLine 0 0 "Cache on device hard disk"}
							7   {WriteWordLine 0 0 "Cache on server persisted"}
							8   {WriteWordLine 0 0 "Cache on device hard drive persisted (NT 6.1 and later)"}
							default {WriteWordLine 0 0 "Cache type could not be determined: $($Disk.writeCacheType)"}
						}
					}
					Else
					{
						switch ($Disk.writeCacheType)
						{
							0   {WriteWordLine 0 0 "Private Image"}
							1   {WriteWordLine 0 0 "Cache on server"}
							2   {WriteWordLine 0 0 "Cache encrypted on server disk" }
							3   {
								WriteWordLine 0 0 "Cache in device RAM"
								WriteWordLine 0 3 "Cache Size: $($Disk.writeCacheSize) MBs"
								}
							4   {WriteWordLine 0 0 "Cache on device's HD"}
							5   {WriteWordLine 0 0 "Cache encrypted on device's HD"}
							6   {WriteWordLine 0 0 "RAM Disk"}
							7   {WriteWordLine 0 0 "Difference Disk"}
							default {WriteWordLine 0 0 "Cache type could not be determined: $($Disk.writeCacheType)"}
						}
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
				write-verbose "Processing Identification Tab"
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
					WriteWordLine 0 3 "Internal name`t: " $Disk.internalName
				}
				If(![String]::IsNullOrEmpty($Disk.originalFile))
				{
					WriteWordLine 0 3 "Original file`t: " $Disk.originalFile
				}
				If(![String]::IsNullOrEmpty($Disk.hardwareTarget))
				{
					WriteWordLine 0 3 "Hardware target: " $Disk.hardwareTarget
				}

				WriteWordLine 0 2 "Microsoft Volume Licensing"
				WriteWordLine 0 3 "Microsoft license type: " -nonewline
				switch ($Disk.licenseMode)
				{
					0 {WriteWordLine 0 0 "None"                          }
					1 {WriteWordLine 0 0 "Multiple Activation Key (MAK)" }
					2 {WriteWordLine 0 0 "Key Management Service (KMS)"  }
					default {WriteWordLine 0 0 "Volume License Mode could not be determined: $($Disk.licenseMode)"}
				}
				#options tab
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
				#PVS 6.x
				WriteWordLine 0 1 "vDisk Properties"
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
				Elseif ($Disk.writeCacheType -eq "7")
				{
					WriteWordLine 0 0 "Difference Disk Image"
				}
				Else
				{
					WriteWordLine 0 0 "Standard Image (multi-device, read-only access)"
					WriteWordLine 0 3 "Cache type`t: " -nonewline
					If($PVSVersion -eq "6")
					{
						switch ($Disk.writeCacheType)
						{
							0   {WriteWordLine 0 0 "Private Image"}
							1   {WriteWordLine 0 0 "Cache on server"}
							3   {WriteWordLine 0 0 "Cache in device RAM"}
							4   {WriteWordLine 0 0 "Cache on device hard disk"}
							7   {WriteWordLine 0 0 "Cache on server persisted"}
							8   {WriteWordLine 0 0 "Cache on device hard drive persisted (NT 6.1 and later)"}
							default {WriteWordLine 0 0 "Cache type could not be determined: $($Disk.writeCacheType)"}
						}
					}
					Else
					{
						switch ($Disk.writeCacheType)
						{
							0   {WriteWordLine 0 0 "Private Image"}
							1   {WriteWordLine 0 0 "Cache on server"}
							2   {WriteWordLine 0 0 "Cache encrypted on server disk" }
							3   {
								WriteWordLine 0 0 "Cache in device RAM"
								WriteWordLine 0 3 "Cache Size: $($Disk.writeCacheSize) MBs"
								}
							4   {WriteWordLine 0 0 "Cache on device's HD"}
							5   {WriteWordLine 0 0 "Cache encrypted on device's HD"}
							6   {WriteWordLine 0 0 "RAM Disk"}
							7   {WriteWordLine 0 0 "Difference Disk"}
							default {WriteWordLine 0 0 "Cache type could not be determined: $($Disk.writeCacheType)"}
						}
					}
				}
				If(![String]::IsNullOrEmpty($Disk.menuText))
				{
					WriteWordLine 0 3 "BIOS menu text`t: " $Disk.menuText
				}
				WriteWordLine 0 3 "AD machine account password management`t: " -nonewline
				If($Disk.adPasswordEnabled -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				
				WriteWordLine 0 3 "Printer management`t`t`t`t: " -nonewline
				If($Disk.printerManagementEnabled -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 3 "Enable streaming of this vDisk`t`t`t: " -nonewline
				If($Disk.Enabled -eq "1")
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				
				write-verbose "Processing Identification Tab"
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
					WriteWordLine 0 3 "Internal name`t: " $Disk.internalName
				}
				If(![String]::IsNullOrEmpty($Disk.originalFile))
				{
					WriteWordLine 0 3 "Original file`t: " $Disk.originalFile
				}
				If(![String]::IsNullOrEmpty($Disk.hardwareTarget))
				{
					WriteWordLine 0 3 "Hardware target: " $Disk.hardwareTarget
				}

				WriteWordLine 0 2 "Microsoft Volume Licensing"
				WriteWordLine 0 3 "Microsoft license type: " -nonewline
				switch ($Disk.licenseMode)
				{
					0 {WriteWordLine 0 0 "None"                          }
					1 {WriteWordLine 0 0 "Multiple Activation Key (MAK)" }
					2 {WriteWordLine 0 0 "Key Management Service (KMS)"  }
					default {WriteWordLine 0 0 "Volume License Mode could not be determined: $($Disk.licenseMode)"}
				}

				write-verbose "Processing Auto Update Tab"
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
				WriteWordLine 0 3 "Major #: " $Disk.majorRelease
				WriteWordLine 0 3 "Minor #: " $Disk.minorRelease
				WriteWordLine 0 3 "Build #`t: " $Disk.build
				WriteWordLine 0 3 "Serial #`t: " $Disk.serialNumber
				
				#process vDisk Load Balancing Menu
				write-verbose "Processing vDisk Load Balancing Menu"
				WriteWordLine 3 1 "vDisk Load Balancing"
				If(![String]::IsNullOrEmpty($Disk.serverName))
				{
					WriteWordLine 0 2 "Use this server to provide the vDisk: " $Disk.serverName
				}
				Else
				{
					WriteWordLine 0 2 "Subnet Affinity`t`t: " -nonewline
					switch ($Disk.subnetAffinity)
					{
						0 {WriteWordLine 0 0 "None"}
						1 {WriteWordLine 0 0 "Best Effort"}
						2 {WriteWordLine 0 0 "Fixed"}
						Default {WriteWordLine 0 0 "Subnet Affinity could not be determined: $($Disk.subnetAffinity)"}
					}
					WriteWordLine 0 1 "Rebalance Enabled`t: " -nonewline
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

	#process all vDisk Update Management in site (PVS 6.x only)
	If($PVSVersion -eq "6")
	{
		write-verbose "Processing vDisk Update Management"
		$Temp = $PVSSite.SiteName
		$GetWhat = "UpdateTask"
		$GetParam = "siteName=$Temp"
		$ErrorTxt = "vDisk Update Management information"
		$Tasks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		WriteWordLine 2 0 " vDisk Update Management"
		If($Tasks -ne $null)
		{
			#process all virtual hosts for this site
			write-verbose "Processing virtual hosts"
			WriteWordLine 0 1 "Hosts"
			$Temp = $PVSSite.SiteName
			$GetWhat = "VirtualHostingPool"
			$GetParam = "siteName=$Temp"
			$ErrorTxt = "Virtual Hosting Pool information"
			$vHosts = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			If($vHosts -ne $null)
			{
				WriteWordLine 3 0 "Hosts"
				ForEach($vHost in $vHosts)
				{
					Write-verbose "Processing virtual host $($vHost.virtualHostingPoolName)"
					write-verbose "Processing General Tab"
					WriteWordLine 4 0 $vHost.virtualHostingPoolName
					WriteWordLine 0 2 "General"
					WriteWordLine 0 3 "Type`t`t: " -nonewline
					switch ($vHost.type)
					{
						0 {WriteWordLine 0 0 "Citrix XenServer"}
						1 {WriteWordLine 0 0 "Microsoft SCVMM/Hyper-V"}
						2 {WriteWordLine 0 0 "VMware vSphere/ESX"}
						Default {WriteWordLine 0 0 "Virtualization Host type could not be determined: $($vHost.type)"}
					}
					WriteWordLine 0 3 "Name`t`t: " $vHost.virtualHostingPoolName
					If(![String]::IsNullOrEmpty($vHost.description))
					{
						WriteWordLine 0 3 "Description`t: " $vHost.description
					}
					WriteWordLine 0 3 "Host`t`t: " $vHost.server
					
					write-verbose "Processing Advanced Tab"
					WriteWordLine 0 2 "Advanced"
					WriteWordLine 0 3 "Update limit`t`t: " $vHost.updateLimit
					WriteWordLine 0 3 "Update timeout`t`t: $($vHost.updateTimeout) minutes"
					WriteWordLine 0 3 "Shutdown timeout`t: $($vHost.shutdownTimeout) minutes"
					WriteWordLine 0 3 "Port`t`t`t: " $vHost.port
				}
			}
			
			WriteWordLine 0 1 "vDisks"
			#process all the Update Managed vDisks for this site
			write-verbose "Processing all Update Managed vDisks for this site"
			$Temp = $PVSSite.SiteName
			$GetParam = "siteName=$Temp"
			$GetWhat = "diskUpdateDevice"
			$ErrorTxt = "Update Managed vDisk information"
			$ManagedvDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			If($ManagedvDisks -ne $null)
			{
				WriteWordLine 3 0 "vDisks"
				ForEach($ManagedvDisk in $ManagedvDisks)
				{
					write-verbose "Processing Managed vDisk $($ManagedvDisk.storeName)`\$($ManagedvDisk.disklocatorName)"
					write-verbose "Processing General Tab"
					WriteWordLine 4 0 "$($ManagedvDisk.storeName)`\$($ManagedvDisk.disklocatorName)"
					WriteWordLine 0 2 "General"
					WriteWordLine 0 3 "vDisk`t`t: " "$($ManagedvDisk.storeName)`\$($ManagedvDisk.disklocatorName)"
					WriteWordLine 0 3 "Virtual Host Connection: " 
					WriteWordLine 0 4 $ManagedvDisk.virtualHostingPoolName
					WriteWordLine 0 3 "VM Name`t: " $ManagedvDisk.deviceName
					WriteWordLine 0 3 "VM MAC`t: " $ManagedvDisk.deviceMac
					WriteWordLine 0 3 "VM Port`t: " $ManagedvDisk.port
									
					write-verbose "Processing Personality Tab"
					#process all personality strings for this device
					$Temp = $ManagedvDisk.deviceName
					$GetWhat = "DevicePersonality"
					$GetParam = "deviceName=$Temp"
					$ErrorTxt = "Device Personality Strings information"
					$PersonalityStrings = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($PersonalityStrings -ne $null)
					{
						WriteWordLine 0 2 "Personality"
						ForEach($PersonalityString in $PersonalityStrings)
						{
							WriteWordLine 0 3 "Name: " $PersonalityString.Name
							WriteWordLine 0 3 "String: " $PersonalityString.Value
						}
					}
					
					write-verbose "Processing Status Tab"
					WriteWordLine 0 2 "Status"
					$Temp = $ManagedvDisk.deviceId
					$GetWhat = "deviceInfo"
					$GetParam = "deviceId=$Temp"
					$ErrorTxt = "Device Info information"
					$Device = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					DeviceStatus $Device
					
					
					write-verbose "Processing Logging Tab"
					WriteWordLine 0 2 "Logging"
					WriteWordLine 0 3 "Logging level: " -nonewline
					switch ($ManagedvDisk.logLevel)
					{
						0   {WriteWordLine 0 0 "Off"    }
						1   {WriteWordLine 0 0 "Fatal"  }
						2   {WriteWordLine 0 0 "Error"  }
						3   {WriteWordLine 0 0 "Warning"}
						4   {WriteWordLine 0 0 "Info"   }
						5   {WriteWordLine 0 0 "Debug"  }
						6   {WriteWordLine 0 0 "Trace"  }
						default {WriteWordLine 0 0 "Logging level could not be determined: $($Server.logLevel)"}
					}
				}
			}
			
			If($Tasks -ne $null)
			{
				ForEach($Task in $Tasks)
				{
					write-verbose "Processing Task $($Task.updateTaskName)"
					write-verbose "Processing General Tab"
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
					write-verbose "Processing Schedule Tab"
					WriteWordLine 0 2 "Schedule"
					WriteWordLine 0 3 "Recurrence: " -nonewline
					switch ($Task.recurrence)
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
						For( $i = 1; $i -le 128; $i = $i * 2 )
						{
							If( ( $Task.dayMask -band $i ) -ne 0 )
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
						switch($Task.monthlyOffset)
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
						For( $i = 1; $i -le 128; $i = $i * 2 )
						{
							If( ( $Task.dayMask -band $i ) -ne 0 )
							{
								WriteWordLine 0 0 $dayMask.$i
							}
						}
					}
					
					write-verbose "Processing vDisks Tab"
					WriteWordLine 0 2 "vDisks"
					WriteWordLine 0 3 "vDisks to be updated by this task:"
					$Temp = $ManagedvDisk.deviceId
					$GetWhat = "diskUpdateDevice"
					$GetParam = "deviceId=$Temp"
					$ErrorTxt = "Device Info information"
					$vDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($vDisks -ne $null)
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
					
					write-verbose "Processing ESD Tab"
					WriteWordLine 0 2 "ESD"
					WriteWordLine 0 3 "ESD client to use: " -nonewline
					switch($Task.esdType)
					{
						""     {WriteWordLine 0 0 "None (runs a custom script on the client)"}
						"WSUS" {WriteWordLine 0 0 "Microsoft Windows Update Service (WSUS)"}
						"SCCM" {WriteWordLine 0 0 "Microsoft System Center Configuration Manager (SCCM)"}
						Default {WriteWordLine 0 0 "ESD Client could not be determined: $($Task.esdType)"}
					}
					
					write-verbose "Processing Scripts Tab"
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
					
					write-verbose "Processing Access Tab"
					WriteWordLine 0 2 "Access"
					WriteWordLine 0 3 "Upon successful completion, access assigned to the vDisk: " -nonewline
					switch($Task.postUpdateApprove)
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
	write-verbose "Processing all device collections in site"
	$Temp = $PVSSite.SiteName
	$GetWhat = "Collection"
	$GetParam = "siteName=$Temp"
	$ErrorTxt = "Device Collection information"
	$Collections = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	If($Collections -ne $null)
	{
		WriteWordLine 2 0 "Device Collections"
		ForEach($Collection in $Collections)
		{
			write-verbose "Processing Collection $($Collection.collectionName)"
			write-verbose "Processing General Tab"
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

			write-verbose "Processing Security Tab"
			WriteWordLine 0 2 "Security"
			$Temp = $Collection.collectionId
			$GetWhat = "authGroup"
			$GetParam = "collectionId=$Temp"
			$ErrorTxt = "Device Collection information"
			$AuthGroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

			$DeviceAdmins = $False
			If($AuthGroups -ne $null)
			{
				WriteWordLine 0 3 "Groups with 'Device Administrator' access:"
				ForEach($AuthGroup in $AuthGroups)
				{
					$Temp = $authgroup.authGroupName
					$GetWhat = "authgroupusage"
					$GetParam = "authgroupname=$Temp"
					$ErrorTxt = "Device Collection Administrator usage information"
					$AuthGroupUsages = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($AuthGroupUsages -ne $null)
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
			If($AuthGroups -ne $null)
			{
				WriteWordLine 0 3 "Groups with 'Device Operator' access:"
				ForEach($AuthGroup in $AuthGroups)
				{
					$Temp = $authgroup.authGroupName
					$GetWhat = "authgroupusage"
					$GetParam = "authgroupname=$Temp"
					$ErrorTxt = "Device Collection Operator usage information"
					$AuthGroupUsages = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($AuthGroupUsages -ne $null)
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

			write-verbose "Processing Auto-Add Tab"
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
			write-verbose "Processing each collection process for each device"
			$Temp = $Collection.collectionId
			$GetWhat = "deviceInfo"
			$GetParam = "collectionId=$Temp"
			$ErrorTxt = "Device Info information"
			$Devices = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			
			If($Devices -ne $null)
			{
				ForEach($Device in $Devices)
				{
					write-verbose "Processing Device $($Device.deviceName)"
					If($Device.type -eq "3")
					{
						WriteWordLine 0 1 "Device with Personal vDisk Properties"
					}
					Else
					{
						WriteWordLine 0 1 "Target Device Properties"
					}
					write-verbose "Processing General Tab"
					WriteWordLine 0 2 "General"
					WriteWordLine 0 3 "Name`t`t`t: " $Device.deviceName
					If(![String]::IsNullOrEmpty($Device.description))
					{
						WriteWordLine 0 3 "Description`t`t: " $Device.description
					}
					If($PVSVersion -eq "6" -and $Device.type -ne "3")
					{
						WriteWordLine 0 3 "Type`t`t`t: " -nonewline
						switch ($Device.type)
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
						switch ($Device.bootFrom)
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
					write-verbose "Processing vDisks Tab"
					WriteWordLine 0 2 "vDisks"
					#process all vdisks for this device
					$Temp = $Device.deviceName
					$GetWhat = "DiskInfo"
					$GetParam = "deviceName=$Temp"
					$ErrorTxt = "Device vDisk information"
					$vDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($vDisks -ne $null)
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
					write-verbose "Processing all bootstrap files for this device"
					$Temp = $Device.deviceName
					$GetWhat = "DeviceBootstraps"
					$GetParam = "deviceName=$Temp"
					$ErrorTxt = "Device Bootstrap information"
					$Bootstraps = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($Bootstraps -ne $null)
					{
						ForEach($Bootstrap in $Bootstraps)
						{
							WriteWordLine 0 4 "Custom bootstrap file: " -nonewline
							WriteWordLine 0 0 "$($Bootstrap.bootstrap) `($($Bootstrap.menuText)`)"
						}
					}
					
					write-verbose "Processing Authentication Tab"
					WriteWordLine 0 2 "Authentication"
					WriteWordLine 0 3 "Type of authentication to use for this device: " -nonewline
					switch($Device.authentication)
					{
						0 {WriteWordLine 0 0 "None"}
						1 {WriteWordLine 0 0 "Username and password"; WriteWordLine 0 4 "Username: " $Device.user; WriteWordLine 0 4 "Password: " $Device.password}
						2 {WriteWordLine 0 0 "External verification (User supplied method)"}
						Default {WriteWordLine 0 0 "Authentication type could not be determined: $($Device.authentication)"}
					}
					
					write-verbose "Processing Personality Tab"
					#process all personality strings for this device
					$Temp = $Device.deviceName
					$GetWhat = "DevicePersonality"
					$GetParam = "deviceName=$Temp"
					$ErrorTxt = "Device Personality Strings information"
					$PersonalityStrings = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($PersonalityStrings -ne $null)
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
		write-verbose "Processing all user groups in site"
		$Temp = $PVSSite.siteName
		$GetWhat = "UserGroup"
		$GetParam = "siteName=$Temp"
		$ErrorTxt = "User Group information"
		$UserGroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		WriteWordLine 2 0 "User Group Properties"
		If($UserGroups -ne $null)
		{
			ForEach($UserGroup in $UserGroups)
			{
				write-verbose "Processing User Group $($UserGroup.userGroupName)"
				write-verbose "Processing General Tab"
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
				write-verbose "Process all vDisks for user group"
				$Temp = $UserGroup.userGroupId
				$GetWhat = "DiskInfo"
				$GetParam = "userGroupId=$Temp"
				$ErrorTxt = "User Group Disk information"
				$vDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt

				write-verbose "Processing vDisk Tab"
				WriteWordLine 0 1 "vDisk"
				WriteWordLine 0 2 "vDisks for this user group:"
				If($vDisks -ne $null)
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
	write-verbose "Processing all site views in site"
	$Temp = $PVSSite.siteName
	$GetWhat = "SiteView"
	$GetParam = "siteName=$Temp"
	$ErrorTxt = "Site View information"
	$SiteViews = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	
	WriteWordLine 2 0 "Site Views"
	If($SiteViews -ne $null)
	{
		ForEach($SiteView in $SiteViews)
		{
			write-verbose "Processing Site View $($SiteView.siteViewName)"
			write-verbose "Processing General Tab"
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
			
			write-verbose "Processing Members Tab"
			WriteWordLine 0 2 "Members"
			#process each target device contained in the site view
			$Temp = $SiteView.siteViewId
			$GetWhat = "Device"
			$GetParam = "siteViewId=$Temp"
			$ErrorTxt = "Site View Device Members information"
			$Members = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			If($Members -ne $null)
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
}

$PVSSites            = $null
$authgroups          = $null
$servers             = $null
$stores              = $null
$bootstrapnames      = $null
$tempserverbootstrap = $null
$serverbootstraps    = $null
$UserGroups          = $null
$Disks               = $null
$vDisks              = $null
$Members             = $null
$SiteViews           = $null

#process the farm views now
write-verbose "Processing all PVS Farm Views"
$selection.InsertNewPage()
WriteWordLine 1 0 "Farm Views"
$Temp = $PVSSite.siteName
$GetWhat = "FarmView"
$GetParam = ""
$ErrorTxt = "Farm View information"
$FarmViews = BuildPVSObject $GetWhat $GetParam $ErrorTxt
If($FarmViews -ne $null)
{
	ForEach($FarmView in $FarmViews)
	{
		write-verbose "Processing Farm View $($FarmView.farmViewName)"
		write-verbose "Processing General Tab"
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
		
		write-verbose "Processing Members Tab"
		WriteWordLine 0 2 "Members"
		#process each target device contained in the farm view
		$Temp = $FarmView.farmViewId
		$GetWhat = "Device"
		$GetParam = "farmViewId=$Temp"
		$ErrorTxt = "Farm View Device Members information"
		$Members = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		If($Members -ne $null)
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
$FarmViews = $null
$Members = $null

#process the stores now
write-verbose "Processing Stores"
$selection.InsertNewPage()
WriteWordLine 1 0 "Stores Properties"
$GetWhat = "Store"
$GetParam = ""
$ErrorTxt = "Farm Store information"
$Stores = BuildPVSObject $GetWhat $GetParam $ErrorTxt
If($Stores -ne $null)
{
	ForEach($Store in $Stores)
	{
		write-verbose "Processing Store $($Store.StoreName)"
		write-verbose "Processing General Tab"
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
		
		write-verbose "Processing Servers Tab"
		WriteWordLine 0 1 "Servers"
		#find the servers (and the site) that serve this store
		$GetWhat = "Server"
		$GetParam = ""
		$ErrorTxt = "Server information"
		$Servers = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		$StoreSite = ""
		$StoreServers = @()
		If($Servers -ne $null)
		{
			ForEach($Server in $Servers)
			{
				write-verbose "Processing Server $($Server.serverName)"
				$Temp = $Server.serverName
				$GetWhat = "ServerStore"
				$GetParam = "serverName=$Temp"
				$ErrorTxt = "Server Store information"
				$ServerStore = BuildPVSObject $GetWhat $GetParam $ErrorTxt
				If($ServerStore -ne $null -and $ServerStore.storeName -eq $Store.StoreName)
				{
					$StoreSite = $Server.siteName
					$StoreServers += $Server.serverName
				}
			}	
		}
		WriteWordLine 0 2 "Site: " $StoreSite
		WriteWordLine 0 2 "Servers that provide this store:"
		ForEach($StoreServer in $StoreServers)
		{
			WriteWordLine 0 3 $StoreServer
		}

		write-verbose "Processing Paths Tab"
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

$Stores = $null
$Servers = $null
$StoreSite = $null
$StoreServers = $null
$ServerStore = $null

write-verbose "Finishing up Word document"
#end of document processing
#Update document properties

If($CoverPagesExist)
{
	write-verbose "Set Cover Page Properties"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Company" $CompanyName
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Title" $title
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Subject" "Provisioning Services $PVSFullVersion Farm Inventory"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Author" $username

	#Get the Coverpage XML part
	$cp=$doc.CustomXMLParts | where {$_.NamespaceURI -match "coverPageProps$"}

	#get the abstract XML part
	$ab=$cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}
	#set the text
	[string]$abstract="Citrix Provisioning Services Inventory for $CompanyName"
	$ab.Text=$abstract

	$ab=$cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
	#set the text
	[string]$abstract=( Get-Date -Format d ).ToString()
	$ab.Text=$abstract

	write-verbose "Update the Table of Contents"
	#update the Table of Contents
	$doc.TablesOfContents.item(1).Update()
}

write-verbose "Save and Close document and Shutdown Word"
If ($WordVersion -eq 12)
{
	#Word 2007
	$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
	$doc.SaveAs($filename, $SaveFormat)
}
Else
{
	#the $saveFormat below passes StrictMode 2
	#I found this at the following two links
	#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
	#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
	$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
	$doc.SaveAs([REF]$filename, [ref]$SaveFormat)
}

$doc.Close()
$Word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
Remove-Variable -Name word
[gc]::collect() 
[gc]::WaitForPendingFinalizers()