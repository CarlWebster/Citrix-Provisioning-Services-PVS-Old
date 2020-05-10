#Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#This script written for "Benji", March 19, 2012
#Thanks to Michael B. Smith, Joe Shonk and Stephane Thirion
#for testing and fine-tuning tips 

Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com
{
	Param( [int]$tabs = 0, [string]$name = ’’, [string]$value = ’’, [string]$newline = “`n”, [switch]$nonewline )

	While( $tabs –gt 0 ) { $global:output += “`t”; $tabs--; }

	If( $nonewline )
	{
		$global:output += $name + $value
	}
	Else
	{
		$global:output += $name + $value + $newline
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

			$index = $record.IndexOf( ‘:’ )
			if( $index –gt 0 )
			{
				$property = $record.SubString( 0, $index  )
				$value    = $record.SubString( $index + 2 )
				If($property -ne "Executing")
				{
					Add-Member –inputObject $SingleObject –MemberType NoteProperty –Name $property –Value $value
				}
			}
		}
		$PluralObject += $SingleObject
		Return $PluralObject
	}
	Else 
	{
		line 0 "$($TextForErrorMsg) could not be retrieved"
		line 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
	}
}

Function DeviceStatus
{
	Param( $xDevice )

	If($xDevice.active -eq "0" -or $xDevice -eq $null)
	{
		line 2 "Target device inactive"
	}
	Else
	{
		line 2 "Target device active"
		line 2 "IP Address: " $xDevice.ip
		line 2 "Server: " -nonewline
		line 0 "$($xDevice.serverName) `($($xDevice.serverIpConnection)`: $($xDevice.serverPortConnection)`)"
		line 2 "Retries: " $xDevice.status
		line 2 "vDisk: " $xDevice.diskLocatorName
		line 2 "vDisk version: " $xDevice.diskVersion
		line 2 "vDisk full name: " $xDevice.diskFileName
		line 2 "vDisk access: " -nonewline
		switch ($xDevice.diskVersionAccess)
		{
			0 {line 0 "Production"}
			1 {line 0 "Test"}
			2 {line 0 "Maintenance"}
			3 {line 0 "Personal vDisk"}
			Default {line 0 "vDisk access type could not be determined: $($xDevice.diskVersionAccess)"}
		}
		switch($xDevice.licenseType)
		{
			0 {line 2 "No License"}
			1 {line 2 "Desktop License"}
			2 {line 2 "Server License"}
			5 {line 2 "OEM SmartClient License"}
			6 {line 2 "XenApp License"}
			7 {line 2 "XenDesktop License"}
			Default {line 2 "Device license type could not be determined: $($xDevice.licenseType)"}
		}
		
		line 1 "  Logging tab"
		line 2 "Logging level: " -nonewline
		switch ($xDevice.logLevel)
		{
			0   {line 0 "Off"    }
			1   {line 0 "Fatal"  }
			2   {line 0 "Error"  }
			3   {line 0 "Warning"}
			4   {line 0 "Info"   }
			5   {line 0 "Debug"  }
			6   {line 0 "Trace"  }
			default {line 0 "Logging level could not be determined: $($xDevice.logLevel)"}
		}
		line 2 ""
	}
}

$global:output = ""

#get PVS major version
$PVSVersion = ""

$error.Clear()
$tempversion = mcli-info version
If( $error.Count -eq 0 )
{
	#build PVS version values
	$version = new-object System.Object 
	foreach( $record in $tempversion )
	{
		$index = $record.IndexOf( ‘:’ )
		if( $index –gt 0 )
		{
			$property = $record.SubString( 0, $index)
			$value = $record.SubString( $index + 2 )
			Add-Member –inputObject $version –MemberType NoteProperty –Name $property –Value $value
		}
	}
} 
Else 
{
	line 0 "PVS version information could not be retrieved"
	line 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
	line 0 "Script is terminating"
	#without version info, script should not proceed
	Break
}

$PVSVersion = $Version.mapiVersion.SubString(0,1)
$tempversion = $null
$version = $null

$FarmAutoAddEnabled = $false

#build PVS farm values
#there can only be one farm
$GetWhat = "Farm"
$GetParam = ""
$ErrorTxt = "PVS Farm information"
$farm = BuildPVSObject $GetWhat $GetParam $ErrorTxt

If($Farm -eq $null)
{
	#without farm info, script should not proceed
	Break
}

line 0 "PVS Farm Information"
#general tab
line 1 "General tab"
line 2 "Name: " $farm.farmName
line 2 "Description: " $farm.description

#security tab
line 1 "Security tab"
line 2 "Groups with Farm Administrator access:"
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
			line 2 $Group.authGroupName
		}
	}
}

#groups tab
line 1 "Groups tab"
line 2 "All the Security Groups that can be assigned acess rights:"
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
			line 3 $Group.authGroupName
		}
	}
}

#licensing tab
line 1 "Licensing tab"
line 2 "License server name: " $farm.licenseServer
line 2 "License server port: " $farm.licenseServerPort
If($PVSVersion -eq "5")
{
	line 2 "User Datacenter licenses for desktops if no Desktop licenses are available: " -nonewline
	If($farm.licenseTradeUp -eq "1")
	{
		line 0 "Enabled"
	}
	Else
	{
		line 0 "Disabled"
	}
}

#options tab
line 1 "Options tab"
line 2 "Auto-Add"
line 3 "Enable auto-add: " -nonewline
If($farm.autoAddEnabled -eq "1")
{
	line 0 "Enabled"
	line 4 "Add new devices to this site: " $farm.defaultSiteName
	$FarmAutoAddEnabled = $true
}
Else
{
	line 0 "Disabled"	
	$FarmAutoAddEnabled = $false
}
line 2 "Auditing"
line 3 "Enable auditing: " -nonewline
If($farm.auditingEnabled -eq "1")
{
	line 0 "Enabled"
}
Else
{
	line 0 "Disabled"
}
line 2 "Offline database support"
line 3 "Enable offline database support: " -nonewline
If($farm.offlineDatabaseSupportEnabled -eq "1")
{
	line 0 "Enabled"	
}
Else
{
	line 0 "Disabled"
}

If($PVSVersion -eq "6")
{
	#vDisk Version tab
	line 1 "vDisk Version tab"
	line 2 "Alert if number of version from base image exceeds: " $farm.maxVersions
	line 2 "Merge after automated vDisk update, if over alert threshold: " -nonewline
	If($farm.automaticMergeEnabled -eq "1")
	{
		line 0 "Enabled"
	}
	Else
	{
		line 0 "Disabled"
	}
	line 2 "Default access mode for new merge versions: " -nonewline
	switch ($Farm.mergeMode)
	{
		0   {line 0 "Production" }
		1   {line 0 "Test"       }
		2   {line 0 "Maintenance"}
		default {line 0 "Default access mode could not be determined: $($Farm.mergeMode)"}
	}
}

#status tab
line 1 "Status tab"
line 2 "Current status of the farm:"
line 3 "Database server: " $farm.databaseServerName
line 3 "Database instance: " $farm.databaseInstanceName
line 3 "Database: " $farm.databaseName
line 3 "Failover Partner Server: " $farm.failoverPartnerServerName
line 3 "Failover Partner Instance: " $farm.failoverPartnerInstanceName
If($Farm.adGroupsEnabled -eq "1")
{
	line 3 "Active Directory groups are used for access rights"
}
Else
{
	line 3 "Active Directory groups are not used for access rights"
}
	
write-output $global:output
$global:output = ""
$farm = $null
$authgroups = $null

#build site values
$GetWhat = "site"
$GetParam = ""
$ErrorTxt = "PVS Site information"
$PVSSites = BuildPVSObject $GetWhat $GetParam $ErrorTxt

ForEach($PVSSite in $PVSSites)
{
	line 0 "Site properties"
	#general tab
	line 1 "General tab"
	line 2 "Name: " $PVSSite.siteName
	line 2 "Description: " $PVSSite.description

	#security tab
	$temp = $PVSSite.SiteName
	$GetWhat = "authgroup"
	$GetParam = "sitename=$temp"
	$ErrorTxt = "Groups with Site Administrator access"
	$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	line 1 "Security tab"
	line 2 "Groups with Site Administrator access:"
	If($authGroups -ne $null)
	{
		ForEach($Group in $authgroups)
		{
			line 3 $Group.authGroupName
		}
	}
	Else
	{
		line 3 "There are no Site Administrators defined"
	}

	If($PVSVersion -eq "5")
	{
		#MAK tab
		line 1 "MAK tab"
		line 2 "MAK User: " $PVSSite.makUser
		line 2 "Password: " $PVSSite.makPassword
	}

	#options tab
	line 1 "Options tab"
	line 2 "Auto-Add"
	If($PVSVersion -eq "5" -or ($PVSVersion -eq "6" -and $FarmAutoAddEnabled))
	{
		line 3 "Add new devices to this collection: " -nonewline
		If($PVSSite.defaultCollectionName)
		{
			line 0 $PVSSite.defaultCollectionName
		}
		Else
		{
			line 0 "<No default collection>"
		}
	}
	If($PVSVersion -eq "6")
	{
		line 3 "Seconds between vDisk inventory scans: " $PVSSite.inventoryFilePollingInterval

		#vDisk Update tab
		line 1 "vDisk Update tab"
		line 2 "Enable automatic vDisk updates on this site: " -nonewline
		If($PVSSite.enableDiskUpdate -eq "1")
		{
			line 0 "Enabled"
			line 2 "Select the server to run vDisk updates for this site: " $PVSSite.diskUpdateServerName
		}
		Else
		{
			line 0 "Disabled"
		}
	}

	#process all servers in site
	$temp = $PVSSite.SiteName
	$GetWhat = "server"
	$GetParam = "sitename=$temp"
	$ErrorTxt = "Servers for Site $temp"
	$servers = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	line 1 "Servers"
	ForEach($Server in $Servers)
	{
		#general tab
		line 1 " Server Properties"
		line 1 "  General tab"
		line 2 "Name: " $Server.serverName
		line 2 "Description: " $Server.description
		line 2 "Power Rating: " $Server.powerRating
		line 2 "Log events to the server's Windows Event Log: " -nonewline
		IF($Server.eventLoggingEnabled -eq "1")
		{
			line 0 "Enabled"
		}
		Else
		{
			line 0 "Disabled"
		}
			
		line 1 "  Network tab"
		line 2 "IP addresses:"
		$test = $Server.ip.ToString()
		$test1 = $test.replace(",","`n`t`t`t")
		line 3 $test1
		line 2 "Ports"
		line 3 "First port: " $Server.firstPort
		line 3 "Last port: " $Server.lastPort
			
		line 1 "  Stores tab"
		#process all stores for this server
		$temp = $Server.serverName
		$GetWhat = "serverstore"
		$GetParam = "servername=$temp"
		$ErrorTxt = "Store information for server $temp"
		$stores = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		line 2 "Stores that this server supports:"

		If($Stores -ne $null)
		{
			ForEach($store in $stores)
			{
				line 3 "Store: " $store.storename
				line 3 "Path: " -nonewline
				If($store.path.length -gt 0)
				{
					line 0 $store.path
				}
				Else
				{
					line 0 "<Using the default path from the store>"
				}
				line 3 "Write cache paths: " -nonewline
				If($store.cachePath.length -gt 0)
				{
					line 0 $store.cachePath
				}
				Else
				{
					line 0 "<Using the default path from the store>"
				}
				line 4 ""
			}
		}

		line 1 "  Options tab"
		If($PVSVersion -eq "5")
		{
			line 2 "Enable automatic vDisk updates"
			line 3 "Check for new versions of a vDisk: " -nonewline
			If($Server.autoUpdateEnabled -eq "1")
			{
				line 0 "Enabled"
			}
			Else
			{
				line 0 "Disabled"
			}
			line 3 "Check for incremental updates to a vDisk: " -nonewline
			If($Server.incrementalUpdateEnabled -eq "1")
			{
				line 0 "Enabled"
			}
			Else
			{
				line 0 "Disabled"
			}
			$AMorPM = "AM"
			$tempHour = $Server.autoUpdateHour
			If($Server.autoUpdateHour -ge "1" -and $Server.autoUpdateHour -le "13")
			{
				$AMorPM = "AM"
			}
			Else
			{
				$AMorPM = "PM"
				If($Server.autoUpdateHour -eq "0")
				{
					$tempHour = $tempHour + 12
				}
				Else
				{
					$tempHour = $tempHour - 12
				}
			}
			$tempMinute = ""
			If($Server.autoUpdateMinute.length -lt 2)
			{
				$tempMinute = "0" + $Server.autoUpdateMinute
			}
			line 3 "Check for updates daily at: $($tempHour)`:$($tempMinute) $($AMorPM)"
			
		}
		line 2 "Active directory"
		If($PVSVersion -eq "5")
		{
			line 3 "Enable automatic password support: " -nonewline
			If($Server.adMaxPasswordAgeEnabled -eq "1")
			{
				line 0 "Enabled"
				line 3 "Change computer account password every this number of days: " $Server.adMaxPasswordAge
			}
			Else
			{
				line 0 "Disabled"
			}
		}
		Else
		{
			line 3 "Automate computer account password updates: " -nonewline
			If($Server.adMaxPasswordAgeEnabled -eq "1")
			{
				line 0 "Enabled"
				line 3 "Days between password updates: " $Server.adMaxPasswordAge
			}
			Else
			{
				line 0 "Disabled"
			}
		}
			
		line 1 "  Logging tab"
		line 2 "Logging level: " -nonewline
		switch ($Server.logLevel)
		{
			0   {line 0 "Off"    }
			1   {line 0 "Fatal"  }
			2   {line 0 "Error"  }
			3   {line 0 "Warning"}
			4   {line 0 "Info"   }
			5   {line 0 "Debug"  }
			6   {line 0 "Trace"  }
			default {line 0 "Logging level could not be determined: $($Server.logLevel)"}
		}
		line 2 "File size maximum (MB): " $Server.logFileSizeMax
		line 2 "Backup files maximum: " $Server.logFileBackupCopiesMax
		line 2 ""
		
		#advanced button at the bottom
		line 1 "  Advanced button"
		line 2 " Server tab"
		line 3 "Threads per port: " $Server.threadsPerPort
		line 3 "Buffers per thread: " $Server.buffersPerThread
		line 3 "Server cache timeout (seconds): " $Server.serverCacheTimeout
		line 3 "Local concurrent I/O limit (transactions): " $Server.localConcurrentIoLimit
		line 3 "Remote concurrent I/O limit (transactions): " $Server.remoteConcurrentIoLimit

		line 2 " Network tab"
		line 3 "Ethernet MTU (bytes): " $Server.maxTransmissionUnits
		line 3 "I/O burst size (KB): " $Server.ioBurstSize
		line 3 "Socket communications"
		line 4 "Enable non-blocking I/O for network communications: " -nonewline
		If($Server.nonBlockingIoEnabled -eq "1")
		{
			line 0 "Enabled"
		}
		Else
		{
			line 0 "Disabled"
		}

		line 2 " Pacing tab"
		line 3 "Boot pause seconds: " $Server.bootPauseSeconds
		line 3 "Maximum boot time (seconds): " $Server.maxBootSeconds
		line 3 "Maximum devices booting: " $Server.maxBootDevicesAllowed
		line 3 "vDisk Creation pacing: " $Server.vDiskCreatePacing

		line 2 " Device tab"
		line 3 "License timeout (seconds): " $Server.licenseTimeout

		line 1 ""
	}

	#the properties for the servers have been processed. 
	#now to process the stuff available via a right-click on each server

	#Configure Bootstrap is first

	ForEach($Server in $Servers)
	{
		
		#first get all bootstrap files for the server
		$temp = $server.serverName
		$GetWhat = "ServerBootstrapNames"
		$GetParam = "serverName=$temp"
		$ErrorTxt = "Server Bootstrap Name information"
		$BootstrapNames = BuildPVSObject $GetWhat $GetParam $ErrorTxt

		#Now that the list of bootstrap names has been gathered
		#We have the mandatory parameter to get the bootstrap info
		#there should be at least one bootstrap filename
		line 1 ""
		line 1 "  Configure Bootstrap settings for server " $Server.serverName
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
							Add-Member –inputObject $serverbootstrap –MemberType NoteProperty –Name $property –Value $value
						}
						$index = $record.IndexOf( ‘:’ )
						if( $index –gt 0 )
						{
							$property = $record.SubString( 0, $index)
							$value = $record.SubString( $index + 2 )
							If($property -ne "Executing")
							{
								Add-Member –inputObject $serverbootstrap –MemberType NoteProperty –Name $property –Value $value
							}
						}
					}
					$serverbootstraps += $serverbootstrap
				}
				Else
				{
					line 0 "Server Bootstrap information could not be retrieved"
					line 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
				}
			}
			If($ServerBootstraps -ne $null)
			{
				ForEach($ServerBootstrap in $ServerBootstraps)
				{
					line 1 "   General tab"	
					line 2 "Bootstrap file: " $ServerBootstrap.Bootstrapname
					If($ServerBootstrap.bootserver1_Ip -ne "0.0.0.0")
					{
						line 2 "IP Address: " $ServerBootstrap.bootserver1_Ip
						line 2 "Subnet Mask: " $ServerBootstrap.bootserver1_Netmask
						line 2 "Gateway: " $ServerBootstrap.bootserver1_Gateway
						line 2 "Port: " $ServerBootstrap.bootserver1_Port
					}
					If($ServerBootstrap.bootserver2_Ip -ne "0.0.0.0")
					{
						line 2 "IP Address: " $ServerBootstrap.bootserver2_Ip
						line 2 "Subnet Mask: " $ServerBootstrap.bootserver2_Netmask
						line 2 "Gateway: " $ServerBootstrap.bootserver2_Gateway
						line 2 "Port: " $ServerBootstrap.bootserver2_Port
					}
					If($ServerBootstrap.bootserver3_Ip -ne "0.0.0.0")
					{
						line 2 "IP Address: " $ServerBootstrap.bootserver3_Ip
						line 2 "Subnet Mask: " $ServerBootstrap.bootserver3_Netmask
						line 2 "Gateway: " $ServerBootstrap.bootserver3_Gateway
						line 2 "Port: " $ServerBootstrap.bootserver3_Port
					}
					If($ServerBootstrap.bootserver4_Ip -ne "0.0.0.0")
					{
						line 2 "IP Address: " $ServerBootstrap.bootserver4_Ip
						line 2 "Subnet Mask: " $ServerBootstrap.bootserver4_Netmask
						line 2 "Gateway: " $ServerBootstrap.bootserver4_Gateway
						line 2 "Port: " $ServerBootstrap.bootserver4_Port
					}
					line 1 "   Options tab"
					line 2 "Verbose mode: " -nonewline
					If($ServerBootstrap.verboseMode -eq "1")
					{
						line 0 "Enabled"
					}
					Else
					{
						line 0 "Disabled"
					}
					line 2 "Interrupt safe mode: " -nonewline
					If($ServerBootstrap.interruptSafeMode -eq "1")
					{
						line 0 "Enabled"
					}
					Else
					{
						line 0 "Disabled"
					}
					line 2 "Advanced Memory Support: " -nonewline
					If($ServerBootstrap.paeMode -eq "1")
					{
						line 0 "Enabled"
					}
					Else
					{
						line 0 "Disabled"
					}
					line 2 "Network recovery method"
					If($ServerBootstrap.bootFromHdOnFail -eq "0")
					{
						line 3 "Restore network connection"
					}
					Else
					{
						line 3 "Reboot to Hard Drive after " -nonewline
						line 0 $ServerBootstrap.recoveryTime -nonewline
						line 0 " seconds"
					}
					line 2 "Timeouts"
					line 3 "Login polling timeout (milliseconds): " -nonewline
					If($ServerBootstrap.pollingTimeout -eq "")
					{
						line 0 "5000"
					}
					Else
					{
						line 0 $ServerBootstrap.pollingTimeout
					}
					line 3 "Login general timeout (milliseconds): " -nonewline
					If($ServerBootstrap.generalTimeout -eq "")
					{
						line 0 "5000"
					}
					Else
					{
						line 0 $ServerBootstrap.generalTimeout
					}
				}
			}
		}
		Else
		{
			line 2 "No Bootstrap names available"
		}
	}		

<#	#Configure BIOS Bootstrap is last
	#this section has been commented out as it causes fatal errors when run on VMs
	
	ForEach($Server in $Servers)
	{
		$temp = $server.serverName
		$GetWhat = "ServerBiosBootstrap"
		$GetParam = "serverName=$Temp"
		$ErrorTxt = "Server Bios Bootstrap information"
		$BiosBootstraps = BuildPVSObject $GetWhat $GetParam $ErrorTxt

		line 1 ""
		line 1 "  Configure BIOS Bootstrap settings for server " $Server.serverName

		If($BiosBootstraps -ne $null)
			ForEach($BiosBootstrap in $BiosBootstraps)
			{
				line 1 "   General tab"
				line 2 "Automatically update the BIOS on the target device with these settings: " -nonewline
				If($BiosBootstraps.enabled -eq "1")
				{
					line 0 "Enabled"
				}
				Else
				{
					line 0 "Disabled"
				}

				line 1 "   Target Device IP tab"
				If($BiosBootstraps.dhcpEnabled -eq "1")
				{
					line 2 "Use DHCP to retrieve target device IP"
				}
				Else
				{
					line 2 "Use static target device IP"
					line 3 "Primary DNS: " $BiosBootstraps.dnsIpAddress1
					line 3 "Secondary DNS: " $BiosBootstraps.dnsIpAddress2
					line 3 "Domain name: " $BiosBootstraps.domain
				}

				line 1 "   Server Lookup tab"
				If($BiosBootstraps.lookup -eq "1")
				{
					line 2 "Use DNS to find server"
					line 3 "Host name: " $BiosBootstraps.serverName
				}
				Else
				{
					line 2 "Use specific servers"
					If($BiosBootstraps.bootserver1_Ip -ne "0.0.0.0")
					{
						line 3 "IP Address: " $BiosBootstraps.bootserver1_Ip
						line 3 "Port: " $BiosBootstraps.bootserver1_Port
					}
					If($BiosBootstraps.bootserver2_Ip -ne "0.0.0.0")
					{
						line 3 "IP Address: " $BiosBootstraps.bootserver2_Ip
						line 3 "Port: " $BiosBootstraps.bootserver2_Port
					}
					If($BiosBootstraps.bootserver3_Ip -ne "0.0.0.0")
					{
						line 3 "IP Address: " $BiosBootstraps.bootserver3_Ip
						line 3 "Port: " $BiosBootstraps.bootserver3_Port
					}
					If($BiosBootstraps.bootserver4_Ip -ne "0.0.0.0")
					{
						line 3 "IP Address: " $BiosBootstraps.bootserver4_Ip
						line 3 "Port: " $BiosBootstraps.bootserver4_Port
					}
				}

				line 1 "   Options tab"
				line 2 "Verbose mode: " -nonewline
				If($BiosBootstraps.verboseMode -eq "1")
				{
					line 0 "Enabled"
				}
				Else
				{
					line 0 "Disabled"
				}
				line 2 "Interrupt safe mode: " -nonewline
				If($BiosBootstraps.interruptSafeMode -eq "1")
				{
					line 0 "Enabled"
				}
				Else
				{
					line 0 "Disabled"
				}
				line 2 "Advanced Memory Support: " -nonewline
				If($BiosBootstraps.paeMode -eq "1")
				{
					line 0 "Enabled"
				}
				Else
				{
					line 0 "Disabled"
				}
				line 2 "Network recovery method"
				If($BiosBootstraps.bootFromHdOnFail -eq "0")
				{
					line 3 "Restore network connection"
				}
				Else
				{
					line 3 "Reboot to Hard Drive after " -nonewline
					line 0 $BiosBootstraps.recoveryTime -nonewline
					line 0 " seconds"
				}
				line 2 "Timeouts"
				line 3 "Login polling timeout (milliseconds): " -nonewline
				If($BiosBootstraps.pollingTimeout -eq "")
				{
					line 0 "5000"
				}
				Else
				{
					line 0 $BiosBootstraps.pollingTimeout
				}
				line 3 "Login general timeout (milliseconds): " -nonewline
				If($BiosBootstraps.generalTimeout -eq "")
				{
					line 0 "5000"
				}
				Else
				{
					line 0 $BiosBootstraps.generalTimeout
				}
			}
		}
	}
#>

	#process all vDisks in site
	$Temp = $PVSSite.SiteName
	$GetWhat = "DiskInfo"
	$GetParam = "siteName=$Temp"
	$ErrorTxt = "Disk information"
	$Disks = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	line 1 "vDisk Pool"
	line 1 " vDisk Properties"
	If($Disks -ne $null)
	{
		ForEach($Disk in $Disks)
		{
			line 1 " General tab"
			line 2 "Site: " $Disk.siteName
			line 2 "Store: " $Disk.storeName
			line 2 "Filename: " $Disk.diskLocatorName
			line 2 "Size: " (($Disk.diskSize/1024)/1024)/1024 -nonewline
			line 0 " MB"
			line 2 "VHD block size: " $Disk.vhdBlockSize -nonewline
			line 0 " KB"
			line 2 "Access mode"
			line 3 "Access mode: " -nonewline
			If($Disk.writeCacheType -eq "0")
			{
				line 0 "Private Image (single device, read/write access)"
			}
			Else
			{
				line 0 "Standard Image (multi-device, read-only access)"
			}
			line 3 "Cache type: " -nonewline
			If($PVSVersion -eq "6")
			{
				switch ($Disk.writeCacheType)
				{
					0   {line 0 "Private Image"             }
					1   {line 0 "Cache on server"           }
					3   {line 0 "Cache in device RAM"       }
					4   {line 0 "Cache on device hard disk" }
					7   {line 0 "Cache on server persisted" }
					8   {line 0 "Cache on device hard drive persisted (NT 6.1 and later)"}
					default {line 0 "Cache type could not be determined: $($Disk.writeCacheType)"}
				}
			}
			Else
			{
				switch ($Disk.writeCacheType)
				{
					0   {line 0 "Private Image"             }
					1   {line 0 "Cache on server"           }
					2   {line 0 "Cache on server encrypted" }
					3   {line 0 "RAM"                       }
					4   {line 0 "Hard Disk"                 }
					5   {line 0 "Hard Disk Encrypted"       }
					6   {line 0 "RAM Disk"                  }
					7   {line 0 "Difference Disk"           }
					default {line 0 "Cache type could not be determined: $($Disk.writeCacheType)"}
				}
			}		
			line 2 "BIOS boot menu text: " $Disk.menuText
			line 2 "Enable Active Directory machine account password management: " -nonewline
			If($Disk.adPasswordEnabled -eq "1")
			{
				line 0 "Enabled"
			}
			Else
			{
				line 0 "Disabled"
			}
			
			line 2 "Enable printer management: " -nonewline
			If($Disk.printerManagementEnabled -eq "1")
			{
				line 0 "Enabled"
			}
			Else
			{
				line 0 "Disabled"
			}

			line 2 "Enable streaming of this vDisk: " -nonewline
			If($Disk.Enabled -eq "1")
			{
				line 0 "Enabled"
			}
			Else
			{
				line 0 "Disabled"
			}
		
			line 1 " Identification tab"
			line 2 "Description: " $Disk.longDescription
			line 2 "Date: " $Disk.date
			line 2 "Author: " $Disk.author
			line 2 "Title: " $Disk.title
			line 2 "Company: " $Disk.company
			line 2 "Internal name: " $Disk.internalName
			line 2 "Original file: " $Disk.originalFile
			line 2 "Hardware target: " $Disk.hardwareTarget

			line 1 " Microsoft Volume Licensing tab"
			line 2 "Microsoft license type: " -nonewline
			switch ($Disk.licenseMode)
			{
				0 {line 0 "None"                          }
				1 {line 0 "Multiple Activation Key (MAK)" }
				2 {line 0 "Key Management Service (KMS)"  }
				default {line 0 "License Mode could not be determined: $($Disk.licenseMode)"}
			}

			line 1 " Auto Update tab"
			line 2 "Enable automatic updates for the vDisk: " -nonewline
			If($Disk.autoUpdateEnabled -eq "1")
			{
				line 0 "Enabled"
			}
			Else
			{
				line 0 "Disabled"
			}

			line 2 "Apply vDisk updates as soon as they are detected by the server: " -nonewline
			If($Disk.activationDateEnabled -eq "0")
			{
				line 0 "Enabled"
			}
			Else
			{
				line 0 "Disabled"
			}

			line 2 "Schedule the next vDisk update to occur on: " -nonewline
			If($Disk.activationDateEnabled -eq "1")
			{
				line 0 $Disk.activeDate
			}
			Else
			{
				line 0 "N/A"
			}
			line 2 "Class: " $Disk.class
			line 2 "Type: " $Disk.imageType
			line 2 "Major #: " $Disk.majorRelease
			line 2 "Minor #: " $Disk.minorRelease
			line 2 "Build #: " $Disk.build
			line 2 "Serial #: " $Disk.serialNumber
		}
	}

	#process all vDisk Update Management in site (PVS 6.x only)
	If($PVSVersion -eq "6")
	{
		$Temp = $PVSSite.SiteName
		$GetWhat = "UpdateTask"
		$GetParam = "siteName=$Temp"
		$ErrorTxt = "vDisk Update Management information"
		$Tasks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		line 1 " vDisk Update Management"
		If($Tasks -ne $null)
		{
			ForEach($Task in $Tasks)
			{
				line 1 "  Hosts"
				#process all virtual hosts for this site
				$Temp = $PVSSite.SiteName
				$GetWhat = "VirtualHostingPool"
				$GetParam = "siteName=$Temp"
				$ErrorTxt = "Virtual Hosting Pool information"
				$vHosts = BuildPVSObject $GetWhat $GetParam $ErrorTxt
				If($vHosts -ne $null)
				{
					ForEach($vHost in $vHosts)
					{
						line 1 "   General tab"
						line 2 "Type: " -nonewline
						switch ($vHost.type)
						{
							0 {line 0 "Citrix XenServer"}
							1 {line 0 "Microsoft SCVMM/Hyper-V"}
							2 {line 0 "VMware vSphere/ESX"}
							Default {line 0 "Virtualization Host type could not be determined: $($vHost.type)"}
						}
						line 2 "Name: " $vHost.virtualHostingPoolName
						line 2 "Description: " $vHost.description
						line 2 "Host: " $vHost.server
						
						line 1 "   Credentials tab"
						line 2 "Enter the credentials for connecting to the host:"
						line 3 "Username: " $vHost.userName
						line 3 "Password: " $vHost.password
					
						line 1 "   Advanced tab"
						line 2 "Update limit: " $vHost.updateLimit
						line 2 "Update timeout: " $vHost.updateTimeout
						line 2 "Shutdown timeout: " $vHost.shutdownTimeout
					}
				}
				
				line 1 "  vDisks"
				line 1 "   Managed vDisk Properties"
				#process all the vDisks for this task ID
				$Temp = $Task.updateTaskId
				$GetWhat = "diskUpdateDevice"
				$GetParam = "updateTaskId=$Temp"
				$ErrorTxt = "vDisk information"
				$ManagedvDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
				If($ManagedvDisks -ne $null)
				{
					ForEach($ManagedvDisk in $ManagedvDisks)
					{
						line 1 "    General tab"
						line 2 "ManagedvDisk: " "$($ManagedvDisk.store.name)`\$($ManagedvDisk.disklocatorName)"
						line 2 "Virtual Host Connection: " $ManagedvDisk.virtualHostingPoolName
						line 2 "VM Name: " $ManagedvDisk.deviceName
						line 2 "VM MAC: " $ManagedvDisk.deviceMac
						line 2 "VM Port: " $ManagedvDisk.port
										
						line 1 "    Personality tab"
						#process all personality strings for this device
						$Temp = $ManagedvDisk.deviceName
						$GetWhat = "DevicePersonality"
						$GetParam = "deviceName=$Temp"
						$ErrorTxt = "Device Personality Strings information"
						$PersonalityStrings = BuildPVSObject $GetWhat $GetParam $ErrorTxt
						If($PersonalityStrings -ne $null)
						{
							ForEach($PersonalityString in $PersonalityStrings)
							{
								line 2 "Name: " $PersonalityString.Name
								line 2 "String: " $PersonalityString.Value
							}
						}
						
						line 1 "    Status tab"
						$Temp = $ManagedvDisk.deviceId
						$GetWhat = "deviceInfo"
						$GetParam = "deviceId=$Temp"
						$ErrorTxt = "Device Info information"
						$Device = BuildPVSObject $GetWhat $GetParam $ErrorTxt
						DeviceStatus $Device
					}
				}
				
				line 1 "  Tasks"
				line 1 "   Update Task Properties"
				line 1 "    General tab"
				line 2 "Name: " $Task.updateTaskName
				line 2 "Description: " $Task.description
				line 2 "Disable this task: " -nonewline
				If($Task.enabled -eq "1")
				{
					line 0 "Not checked"
				}
				Else
				{
					line 0 "Checked"
				}
				line 1 "    Schedule tab"
				line 2 "Recurrence: " -nonewline
				switch ($Task.recurrence)
				{
					0 {line 0 "None"}
					1 {line 0 "Daily Everyday"}
					2 {line 0 "Daily Weekdays only"}
					3 {line 0 "Weekly"}
					4 {line 0 "Monthly Date"}
					5 {line 0 "Monthly Type"}
					Default {line 0 "Recurrence type could not be determined: $($Task.recurrence)"}
				}
				If($Task.recurrence -ne "0")
				{
					$AMorPM = "AM"
					$tempHour = $Task.Hour
					If($Task.Hour -ge "1" -and $Task.Hour -le "13")
					{
						$AMorPM = "AM"
					}
					Else
					{
						$AMorPM = "PM"
						If($Task.Hour -eq "0")
						{
							$tempHour = $tempHour + 12
						}
						Else
						{
							$tempHour = $tempHour - 12
						}
					}
					$tempMinute = ""
					If($Task.Minute.length -lt 2)
					{
						$tempMinute = "0" + $Task.Minute
					}
					line 3 "Run the update at $($tempHour)`:$($tempMinute) $($AMorPM)"
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
					For( $i = 1; $i –le 128; $i = $i * 2 )
					{
						If( ( $Task.dayMask –band $i ) –ne 0 )
						{
							line 3 $dayMask.$i
						}
					}
				}
				If($Task.recurrence -eq "4")
				{
					line 3 "On Date " $Task.date
				}
				If($Task.recurrence -eq "5")
				{
					line 3 "On " -nonewline
					switch($Task.monthlyOffset)
					{
						1 {line 0 "First"}
						2 {line 0 "Second"}
						3 {line 0 "Third"}
						4 {line 0 "Fourth"}
						5 {line 0 "Last"}
						Default {line 0 "Monthly Offset could not be determined: $($Task.monthlyOffset)"}
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
					For( $i = 1; $i –le 128; $i = $i * 2 )
					{
						If( ( $Task.dayMask –band $i ) –ne 0 )
						{
							line 3 $dayMask.$i
						}
					}
				}
				
				line 1 "    vDisks tab"
				line 2 "Select the vDisks to be updated by this task:"
				$Temp = $ManagedvDisk.deviceId
				$GetWhat = "diskUpdateDevice"
				$GetParam = "deviceId=$Temp"
				$ErrorTxt = "Device Info information"
				$vDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
				If($vDisks -ne $null)
				{
					ForEach($vDisk in $vDisks)
					{
						line 3 "vDisk: " -nonewline
						line 0 "$($vDisk.storeName)`\$($vDisk.diskLocatorName)"
						line 3 "Host: " $vDisk.virtualHostingPoolName
						line 3 "VM: " $vDisk.deviceName
					}
				}
				
				line 1 "    ESD tab"
				line 2 "Select ESD client to use: " -nonewline
				switch($Task.esdType)
				{
					""     {line 0 "None (runs a custom script on the client)"}
					"WSUS" {line 0 "Microsoft Windows Update Service (WSUS)"}
					"SCCM" {line 0 "Microsoft System Center Configuration Manager (SCCM)"}
					Default {line 0 "ESD Client could not be determined: $($Task.esdType)"}
				}
				
				line 1 "    Scripts tab"
				line 2 "Scripts that should execute with the vDisk update processing:"
				line 3 "Pre-update script: " $Task.preUpdateScript
				line 3 "Pre-startup script: " $Task.preVmScript
				line 3 "Post-shutdown script: " $Task.postVmScript
				line 3 "Post-update script: " $Task.postUpdateScript
				
				line 1 "    Access tab"
				line 2 "Upon successful completion of the update, select the access to asign to the vDisk: " -nonewline
				switch($Task.postUpdateApprove)
				{
					0 {line 0 "Production"}
					1 {line 0 "Test"}
					2 {line 0 "Maintenance"}
					Default {line 0 "Access method for vDisk could not be determined: $($Task.postUpdateApprove)"}
				}
			}
		}
	}

	#process all device collections in site
	$Temp = $PVSSite.SiteName
	$GetWhat = "Collection"
	$GetParam = "siteName=$Temp"
	$ErrorTxt = "Device Collection information"
	$Collections = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	If($Collections -ne $null)
	{
		ForEach($Collection in $Collections)
		{
			line 1 "  Device Collection Properties"
			line 1 "   General tab"
			line 2 "Name: " $Collection.collectionName
			line 2 "Description: " $Collection.description

			line 1 "   Security tab"
			line 2 "Groups with 'Device Administrator' access:"
			$Temp = $Collection.collectionId
			$GetWhat = "authGroup"
			$GetParam = "collectionId=$Temp"
			$ErrorTxt = "Device Collection information"
			$AuthGroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

			$DeviceAdmins = $False
			If($AuthGroups -ne $null)
			{
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
								line 3 $authgroup.authGroupName
							}
						}
					}
				}
			}
			If(!$DeviceAdmins)
			{
				line 3 "There are no device collection administrators defined"
			}

			line 2 "Groups with 'Device Operator' access:"
			$DeviceOperators = $False
			If($AuthGroups -ne $null)
			{
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
								line 3 $authgroup.authGroupName
							}
						}
					}
				}
			}
			If(!$DeviceOperators)
			{
				line 3 "There are no device collection operators defined"
			}

			line 1 "   Auto-Add tab"
			If($FarmAutoAddEnabled)
			{
				line 2 "Template target device: " $Collection.templateDeviceName
				line 2 "Device Name"
				line 3 "Prefix: " $Collection.autoAddPrefix
				line 3 "Length: " $Collection.autoAddNumberLength
				line 3 "Zero fill: " -nonewline
				If($Collection.autoAddZeroFill -eq "1")
				{
					line 0 "Enabled"
				}
				Else
				{
					line 0 "Disabled"
				}
				line 3 "Suffix: " $Collection.autoAddSuffix
				line 3 "Last incremental number: " $Collection.lastAutoAddDeviceNumber
			}
			Else
			{
				line 2 "The auto-add feature is not enabled at the PVS Farm level"
			}
			#for each collection process each device
			$Temp = $Collection.collectionId
			$GetWhat = "deviceInfo"
			$GetParam = "collectionId=$Temp"
			$ErrorTxt = "Device Info information"
			$Devices = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			
			If($Devices -ne $null)
			{
				ForEach($Device in $Devices)
				{
					If($Device.type -eq "3")
					{
						line 1 " Device with Personal vDisk Properties"
					}
					Else
					{
						line 1 " Target Device Properties"
					}
					line 1 "  General tab"
					line 2 "Name: " $Device.deviceName
					line 2 "Description: " $Device.description
					If($PVSVersion -eq "6" -and $Device.type -ne "3")
					{
						line 2 "Type: " -nonewline
						switch ($Device.type)
						{
							0 {line 0 "Production"}
							1 {line 0 "Test"}
							2 {line 0 "Maintenance"}
							3 {line 0 "Personal vDisk"}
							Default {line 0 "Device type could not be determined: $($Device.type)"}
						}
					}
					If($Device.type -ne "3")
					{
						line 2 "Boot from: " -nonewline
						switch ($Device.bootFrom)
						{
							1 {line 0 "vDisk"}
							2 {line 0 "Hard Disk"}
							3 {line 0 "Floppy Disk"}
							Default {line 0 "Boot from could not be determined: $($Device.bootFrom)"}
						}
					}
					line 2 "MAC: " $Device.deviceMac
					line 2 "Port: " $Device.port
					If($Device.type -ne "3")
					{
						line 2 "Class: " $Device.className
						line 2 "Disable this device: " -nonewline
						If($Device.enabled -eq "1")
						{
							line 0 "Unchecked"
						}
						Else
						{
							line 0 "Checked"
						}
					}
					Else
					{
						line 2 "vDisk: " $Device.diskLocatorName
						line 2 "Personal vDisk Drive: " $Device.pvdDriveLetter
					}
					line 1 "  vDisks tab"
					line 2 "vDisks for this Device:"
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
							line 3 "Name: " -nonewline
							line 0 "$($vDisk.storeName)`\$($vDisk.diskLocatorName)"
						}
					}
					line 2 "Options"
					line 3 "List local hard drive in boot menu: " -nonewline
					If($Device.localDiskEnabled -eq "1")
					{
						line 0 "Enabled"
					}
					Else
					{
						line 0 "Disabled"
					}
					#process all bootstrap files for this device
					$Temp = $Device.deviceName
					$GetWhat = "DeviceBootstraps"
					$GetParam = "deviceName=$Temp"
					$ErrorTxt = "Device Bootstrap information"
					$Bootstraps = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($Bootstraps -ne $null)
					{
						ForEach($Bootstrap in $Bootstraps)
						{
							line 3 "Custom bootstrap file: " -nonewline
							line 0 "$($Bootstrap.bootstrap) `($($Bootstrap.menuText)`)"
						}
					}
					
					line 1 "  Authentication tab"
					line 2 "Select the type of authentication to use for this device"
					line 3 "Authentication: " -nonewline
					switch($Device.authentication)
					{
						0 {line 0 "None"}
						1 {line 0 "Username and password"; line 3 "Username: " $Device.user; line 3 "Password: " $Device.password}
						2 {line 0 "External verification (User supplied method)"}
						Default {line 0 "Authentication type could not be determined: $($Device.authentication)"}
					}
					
					line 1 "  Personality tab"
					#process all personality strings for this device
					$Temp = $Device.deviceName
					$GetWhat = "DevicePersonality"
					$GetParam = "deviceName=$Temp"
					$ErrorTxt = "Device Personality Strings information"
					$PersonalityStrings = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					If($PersonalityStrings -ne $null)
					{
						ForEach($PersonalityString in $PersonalityStrings)
						{
							line 2 "Name: " $PersonalityString.Name
							line 2 "String: " $PersonalityString.Value
						}
					}
					
					line 1 "  Status tab"
					DeviceStatus $Device
				}
			}
		}
	}

	#process all user groups in site (PVS 5.6 only)
	If($PVSVersion -eq "5")
	{
		$Temp = $PVSSite.siteName
		$GetWhat = "UserGroup"
		$GetParam = "siteName=$Temp"
		$ErrorTxt = "User Group information"
		$UserGroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		line 1 "User Group Properties"
		If($UserGroups -ne $null)
		{
			ForEach($UserGroup in $UserGroups)
			{
				line 1 " General tab"
				line 2 "Name: " $UserGroup.userGroupName
				line 2 "Description: " $UserGroup.description
				line 2 "Class: " $UserGroup.className
				line 2 "Disable this user group: " -nonewline
				If($UserGroup.enabled -eq "1")
				{
					line 0 "Not Checked"
				}
				Else
				{
					line 0 "Checked"
				}
				#process all vDisks for usergroup
				$Temp = $UserGroup.userGroupId
				$GetWhat = "DiskInfo"
				$GetParam = "userGroupId=$Temp"
				$ErrorTxt = "User Group Disk information"
				$vDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt

				line 1 " vDisk tab"
				line 2 "vDisks for this user group:"
				If($vDisks -ne $null)
				{
					ForEach($vDisk in $vDisks)
					{
						line 3 "$($vDisk.storeName)`\$($vDisk.diskLocatorName)"
					}
				}
			}
		}
	}
	
	#process all site views in site
	$Temp = $PVSSite.siteName
	$GetWhat = "SiteView"
	$GetParam = "siteName=$Temp"
	$ErrorTxt = "Site View information"
	$SiteViews = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	
	line 1 " View Properties"
	If($SiteViews -ne $null)
	{
		ForEach($SiteView in $SiteViews)
		{
			line 1 "  General tab"
			line 2 "Name: " $SiteView.siteViewName
			line 2 "Description: " $SiteView.description
			
			line 1 "  Members tab"
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
					line 2 $Member.deviceName
				}
			}
		}
	}
	Else
	{
		line 2 "There are no Site Views configured"
	}
}

write-output $global:output
$global:output       = ""
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
line 0 ""
line 0 "Farm View Properties"
$Temp = $PVSSite.siteName
$GetWhat = "FarmView"
$GetParam = ""
$ErrorTxt = "Farm View information"
$FarmViews = BuildPVSObject $GetWhat $GetParam $ErrorTxt
If($FarmViews -ne $null)
{
	ForEach($FarmView in $FarmViews)
	{
		line 1 "General tab"
		line 2 "Name: " $FarmView.farmViewName
		line 2 "Description: " $FarmView.description
		
		line 1 "Members tab"
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
				line 2 $Member.deviceName
			}
		}
	}
}
Else
{
	line 1 "There are no Farm Views configured"
}
write-output $global:output
$global:output = ""
$FarmViews = $null
$Members = $null

#process the stores now
line 0 ""
line 0 "Stores Properties"
$GetWhat = "Store"
$GetParam = ""
$ErrorTxt = "Farm Store information"
$Stores = BuildPVSObject $GetWhat $GetParam $ErrorTxt
If($Stores -ne $null)
{
	ForEach($Store in $Stores)
	{
		line 1 "General tab"
		line 2 "Name: " $Store.StoreName
		line 2 "Description: " $Store.description
		line 2 "Site that acts as the owner of this store: " -nonewline
		If($Store.siteName -eq $null -or $Store.siteName -eq "")
		{
			line 0 "<none>"
		}
		Else
		{
			line 0 $Store.siteName
		}
		
		line 1 "Servers tab"
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
		line 2 "Site: " $StoreSite
		line 2 "Servers that provide this store:"
		ForEach($StoreServer in $StoreServers)
		{
			line 3 $StoreServer
		}

		line 1 "Paths tab"
		line 2 "Default store path: " $Store.path
		line 2 "Default write-cache paths: "
		If($Store.cachePath -ne $null)
		{
			$WCPaths = $Store.cachePath.replace(",","`n`t`t`t")
			line 3 $WCPaths		
		}
	}
}
Else
{
	line 1 "There are no Stores configured"
}
write-output $global:output
$global:output = ""
$Stores = $null
$Servers = $null
$StoreSite = $null
$StoreServers = $null
$ServerStore = $null
