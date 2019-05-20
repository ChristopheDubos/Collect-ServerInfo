<#
.SYNOPSIS
Collect-ServerInfo.ps1 - PowerShell script to collect information about Windows servers

.DESCRIPTION 
This PowerShell script runs a series of WMI and other queries to collect information
about Windows servers.

.OUTPUTS
Each server's results are output to HTML.

.PARAMETER -Verbose
See more detailed progress as the script is running.

.EXAMPLE
.\Collect-ServerInfo.ps1 SERVER1
Collect information about a single server.

.EXAMPLE
"SERVER1","SERVER2","SERVER3" | .\Collect-ServerInfo.ps1
Collect information about multiple servers.

.EXAMPLE
Get-ADComputer -Filter {OperatingSystem -Like "Windows Server*"} | %{.\Collect-ServerInfo.ps1 $_.DNSHostName}
Collects information about all servers in Active Directory.

Change Log
V1.00, 20/04/2015 - Initial release
V1.01, 01/05/2015 - Updated with better error handling
V1.02, 29/05/2015 - Updated to collect Services and Processes information 
V1.03, 02/06/2015 - Updated to collect Printers / Drivers / Ports / Shares information and improve RPC/WMI connection errors handling
V1.04, 09/06/2015 - Updated to collect Software information via Regsitry for Windows Server 2003/XP computers
V1.05, 12/06/2015 - Updated to collect Server Roles information for Windows Server 2008 and above
V1.06, 19/06/2015 - Updated to collect System Last Boot information
V1.07, 09/07/2015 - Updated to collect Optional componants for Windows Server 2008 and above
V1.08, 02/10/2016 - Updated to collect Scheduled Tasks information for Windows Server 2003/XP computers
V1.09, 27/10/2016 - Updated to collect IIS information
V1.10, 25/11/2016 - Updated to collect TCP connection for Windows Server 2012 and above
V1.11, 21/09/2017 - Updated to collect additional BIOS Manufacturer and Operating System Service Pack information
V1.12, 14/12/2017 - Updated to collect ODBC Drivers and Sourc information
#>


[CmdletBinding()]

Param (

    [parameter(ValueFromPipeline=$True)]
    [string[]]$ComputerName

)

Begin
{
    #Initialize
    Write-Verbose "Initializing"

}

Process
{

    #---------------------------------------------------------------------
    # Process each ComputerName
    #---------------------------------------------------------------------

    if (!($PSCmdlet.MyInvocation.BoundParameters[“Verbose”].IsPresent))
    {
        Write-Host "Processing $ComputerName"
    }

    Write-Verbose "=====> Processing $ComputerName <====="

    $htmlreport = @()
    $htmlbody = @()
    $htmlfile = "$($ComputerName).html"
    $spacer = "<br />"

    #---------------------------------------------------------------------
    # Do 4 pings and calculate the fastest response time
    # Not using the response time in the report yet so it might be
    # removed later.
    # Exit if the ping fails
    #---------------------------------------------------------------------
    
    try
    {
        $bestping = (Test-Connection -ComputerName $ComputerName -Count 4 -ErrorAction STOP | Sort ResponseTime)[0].ResponseTime
    }
    catch
    {
        Write-Warning $_.Exception.Message
        $bestping = "Unable to connect"
    }

    if ($bestping -eq "Unable to connect")
    {
        if (!($PSCmdlet.MyInvocation.BoundParameters[“Verbose”].IsPresent))
        {
            Write-Host "Unable to connect to $ComputerName"
        }

        "Unable to connect to $ComputerName"
    }
    else
    {

        #---------------------------------------------------------------------
        # Collect computer system information and convert to HTML fragment
        # Exit if WMI/RPC connection fails
        #---------------------------------------------------------------------
    
        Write-Verbose "Collecting computer system information"

        $subhead = "<h3>Computer System Information</h3>"
        $htmlbody += $subhead
    
        try
        {
            $csinfo = Get-WmiObject Win32_ComputerSystem -ComputerName $ComputerName -ErrorAction STOP |
                Select-Object Name,Manufacturer,SystemType,Model,
                            @{Name='Physical Processors';Expression={$_.NumberOfProcessors}},
                            @{Name='Logical Processors';Expression={$_.NumberOfLogicalProcessors}},
                            @{Name='Total Physical Memory (Gb)';Expression={
                                $tpm = $_.TotalPhysicalMemory/1GB;
                                "{0:F0}" -f $tpm
                            }},
                            DnsHostName,Domain
       
            $htmlbody += $csinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
       
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
            $csinfo = "WMI connection failed"
        }

    if ($csinfo -eq "WMI connection failed")
    {
        if (!($PSCmdlet.MyInvocation.BoundParameters[“Verbose”].IsPresent))
        {
            Write-Host "Unable to connect to $ComputerName"
        }

        "Unable to connect to $ComputerName"
    }
    else
    {
    
        #---------------------------------------------------------------------
        # Collect operating system information and convert to HTML fragment
        #---------------------------------------------------------------------
    
        Write-Verbose "Collecting operating system information"

        $subhead = "<h3>Operating System Information</h3>"
        $htmlbody += $subhead
    
        try
        {
            $osinfo = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction STOP | 
                Select-Object @{Name='Operating System';Expression={$_.Caption}},
                            @{Name='Architecture';Expression={$_.OSArchitecture}},
                            Version,ServicePackMajorVersion,Organization,
                            @{Name='Install Date';Expression={
                                $installdate = [datetime]::ParseExact($_.InstallDate.SubString(0,8),"yyyyMMdd",$null);
                                $installdate.ToShortDateString()
                            }},
                            @{Name='Boot Date';Expression={
                                $bootdate = [datetime]::ParseExact($_.LastBootUpTime.SubString(0,8),"yyyyMMdd",$null);
                                $bootdate.ToShortDateString()
                                #$boottime = [datetime]::ParseExact($_.LastBootUpTime.SubString(8,14),"hhMMss",$null);
                                #$boottime.ToShortTimeString()
                            }},
                            WindowsDirectory

            $htmlbody += $osinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect physical memory information and convert to HTML fragment
        #---------------------------------------------------------------------

        Write-Verbose "Collecting physical memory information"

        $subhead = "<h3>Physical Memory Information</h3>"
        $htmlbody += $subhead

        try
        {
            $memorybanks = @()
            $physicalmemoryinfo = @(Get-WmiObject Win32_PhysicalMemory -ComputerName $ComputerName -ErrorAction STOP |
                Select-Object DeviceLocator,Manufacturer,Speed,Capacity)

            foreach ($bank in $physicalmemoryinfo)
            {
                $memObject = New-Object PSObject
                $memObject | Add-Member NoteProperty -Name "Device Locator" -Value $bank.DeviceLocator
                $memObject | Add-Member NoteProperty -Name "Manufacturer" -Value $bank.Manufacturer
                $memObject | Add-Member NoteProperty -Name "Speed" -Value $bank.Speed
                $memObject | Add-Member NoteProperty -Name "Capacity (GB)" -Value ("{0:F0}" -f $bank.Capacity/1GB)

                $memorybanks += $memObject
            }

            $htmlbody += $memorybanks | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect pagefile information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>PageFile Information</h3>"
        $htmlbody += $subhead

        Write-Verbose "Collecting pagefile information"

        try
        {
            $pagefileinfo = Get-WmiObject Win32_PageFileUsage -ComputerName $ComputerName -ErrorAction STOP |
                Select-Object @{Name='Pagefile Name';Expression={$_.Name}},
                            @{Name='Allocated Size (Mb)';Expression={$_.AllocatedBaseSize}}

            $htmlbody += $pagefileinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect BIOS information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>BIOS Information</h3>"
        $htmlbody += $subhead

        Write-Verbose "Collecting BIOS information"

        try
        {
            $biosinfo = Get-WmiObject Win32_Bios -ComputerName $ComputerName -ErrorAction STOP |
                Select-Object Status,Version,Manufacturer,Caption,
                            @{Name='Release Date';Expression={
                                $releasedate = [datetime]::ParseExact($_.ReleaseDate.SubString(0,8),"yyyyMMdd",$null);
                                $releasedate.ToShortDateString()
                            }},
                            @{Name='Serial Number';Expression={$_.SerialNumber}}

            $htmlbody += $biosinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect logical disk information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Logical Disk Information</h3>"
        $htmlbody += $subhead

        Write-Verbose "Collecting logical disk information"

        try
        {
            $diskinfo = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName -ErrorAction STOP | 
                Select-Object DeviceID,FileSystem,VolumeName,
                @{Expression={$_.Size /1Gb -as [int]};Label="Total Size (GB)"},
                @{Expression={$_.Freespace / 1Gb -as [int]};Label="Free Space (GB)"}

            $htmlbody += $diskinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect volume information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Volume Information</h3>"
        $htmlbody += $subhead

        Write-Verbose "Collecting volume information"

        try
        {
            $volinfo = Get-WmiObject Win32_Volume -ComputerName $ComputerName -ErrorAction STOP | 
                Select-Object Label,Name,DeviceID,SystemVolume,
                @{Expression={$_.Capacity /1Gb -as [int]};Label="Total Size (GB)"},
                @{Expression={$_.Freespace / 1Gb -as [int]};Label="Free Space (GB)"}

            $htmlbody += $volinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect network interface information and convert to HTML fragment
        #---------------------------------------------------------------------    

        $subhead = "<h3>Network Interface Information</h3>"
        $htmlbody += $subhead

        Write-Verbose "Collecting network interface information"

        try
        {
            $nics = @()
             $nicinfo = @(Get-WmiObject Win32_NetworkAdapter -ComputerName $ComputerName -ErrorAction STOP |
                Select-Object Name,AdapterType,MACAddress,

                @{Name='ConnectionName';Expression={$_.NetConnectionID}},
                @{Name='Enabled';Expression={$_.NetEnabled}},
                @{Name='Speed';Expression={$_.Speed/1000000}})

            $nwinfo = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -ErrorAction STOP |
                Select-Object Description, DHCPServer,  
                @{Name='IpAddress';Expression={$_.IpAddress -join '; '}},  
                @{Name='IpSubnet';Expression={$_.IpSubnet -join '; '}},  
                @{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}},  
                @{Name='DNSServerSearchOrder';Expression={$_.DNSServerSearchOrder -join '; '}}

            foreach ($nic in $nicinfo)
            {
                $nicObject = New-Object PSObject
                $nicObject | Add-Member NoteProperty -Name "Connection Name" -Value $nic.connectionname
                $nicObject | Add-Member NoteProperty -Name "Adapter Name" -Value $nic.Name
                $nicObject | Add-Member NoteProperty -Name "Type" -Value $nic.AdapterType
                $nicObject | Add-Member NoteProperty -Name "MAC" -Value $nic.MACAddress
                $nicObject | Add-Member NoteProperty -Name "Enabled" -Value $nic.Enabled
                $nicObject | Add-Member NoteProperty -Name "Speed (Mbps)" -Value $nic.Speed
        
                $ipaddress = ($nwinfo | Where {$_.Description -eq $nic.Name}).IpAddress
                $nicObject | Add-Member NoteProperty -Name "IPAddress" -Value $ipaddress

                $nics += $nicObject
            }

            $htmlbody += $nics | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect feature information via Win32_ServerFeature WMI and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Server Features Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting features information"
        
        try
        {
            $feature = Get-WmiObject Win32_ServerFeature -ComputerName $ComputerName -ErrorAction STOP | Select-Object ID,Name | Sort-Object Name
        
            $htmlbody += $feature | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect role information via OC Manager Registry key and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Role Information via OC Manager Registry</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting components information via Registry"
        
        try
        {

	    $Base = New-Object PSObject;
	    $Base | Add-Member Noteproperty Name -Value $Null;
	    $Base | Add-Member Noteproperty Version -Value $Null;
	    $Base | Add-Member Noteproperty Publisher -Value $Null;
	    $Results =  New-Object System.Collections.Generic.List[System.Object];
        
		$Registry = $Null;

		try
        {
        
            $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$ComputerName);
        
        }
		
        catch
        {
        
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer

        
        }
		
		if ($Registry)
        {
			
            $UninstallKeys = $Null;
			$SubKey = $Null;
			$UninstallKeys = $Registry.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Setup\OC Manager\Subcomponents",$False);
			$UninstallKeys.GetSubKeyNames()|%{
			$SubKey = $UninstallKeys.OpenSubKey($_,$False);
			$DisplayName = $SubKey.GetValue("DisplayName");

				if ($DisplayName.Length -gt 0)
                {

					$Entry = $Base | Select-Object *
					$Entry.Name = $DisplayName.Trim();
					$Entry.Version = $SubKey.GetValue("DisplayVersion");
					$Entry.Publisher = $SubKey.GetValue("Publisher");
				
                }
					
                [Void]$Results.Add($Entry);

				}

        }
			

        $htmlbody += $results | ConvertTo-Html -Fragment
        $htmlbody += $spacer 

        }

        catch
        {

            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer

        }


        #---------------------------------------------------------------------
        # Collect software information via Win32_Product WMI and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Software Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting software information"
        
        try
        {
            $software = Get-WmiObject Win32_Product -ComputerName $ComputerName -ErrorAction STOP | Select-Object Name,Version,Vendor | Sort-Object Name
        
            $htmlbody += $software | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect software information via Uninstall Registry key and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Software Information via Uninstall Registry</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting software information via Registry"
        
        try
        {

	    $Base = New-Object PSObject;
	    $Base | Add-Member Noteproperty Name -Value $Null;
	    $Base | Add-Member Noteproperty Version -Value $Null;
	    $Base | Add-Member Noteproperty Publisher -Value $Null;
	    $Results =  New-Object System.Collections.Generic.List[System.Object];
        
		$Registry = $Null;

		try
        {
        
            $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$ComputerName);
        
        }
		
        catch
        {
        
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer

        
        }
		
		if ($Registry)
        {
			
            $UninstallKeys = $Null;
			$SubKey = $Null;
			$UninstallKeys = $Registry.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Uninstall",$False);
			$UninstallKeys.GetSubKeyNames()|%{
			$SubKey = $UninstallKeys.OpenSubKey($_,$False);
			$DisplayName = $SubKey.GetValue("DisplayName");

				if ($DisplayName.Length -gt 0)
                {

					$Entry = $Base | Select-Object *
					$Entry.Name = $DisplayName.Trim();
					$Entry.Version = $SubKey.GetValue("DisplayVersion");
					$Entry.Publisher = $SubKey.GetValue("Publisher");
				
                }
					
                [Void]$Results.Add($Entry);

				}

        }
			
				if ([IntPtr]::Size -eq 8)
                {

                    $UninstallKeysWow6432Node = $Null;
                    $SubKeyWow6432Node = $Null;
                    $UninstallKeysWow6432Node = $Registry.OpenSubKey("Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall",$False);

                    if ($UninstallKeysWow6432Node)
                    {

                        $UninstallKeysWow6432Node.GetSubKeyNames()|%{
                        $SubKeyWow6432Node = $UninstallKeysWow6432Node.OpenSubKey($_,$False);
                        $DisplayName = $SubKeyWow6432Node.GetValue("DisplayName");
                        
                        if ($DisplayName.Length -gt 0)
                        {

                        	$Entry = $Base | Select-Object *
                            $Entry.Name = $DisplayName.Trim(); 
                            $Entry.Version = $SubKeyWow6432Node.GetValue("DisplayVersion");
                            $Entry.Publisher = $SubKeyWow6432Node.GetValue("Publisher"); 
                            [Void]$Results.Add($Entry);

                        }
                        }
                	}
                }

        $htmlbody += $results | ConvertTo-Html -Fragment
        $htmlbody += $spacer 

        }

        catch
        {

            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer

        }


        #---------------------------------------------------------------------
        # Collect services information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Services Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting services information"
        
        try
        {
            $service = Get-WmiObject Win32_Service -ComputerName $ComputerName -ErrorAction STOP | Select-Object Caption,Description,PathName,State | Sort-Object Caption
        
            $htmlbody += $service | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }
       

        #---------------------------------------------------------------------
        # Collect Running process information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Process Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting process information"
        
        try
        {
            $process = Get-WmiObject Win32_Process -ComputerName $ComputerName -ErrorAction STOP | Select-Object Caption,Description,ExecutablePath,CommandLine,ProcessID | Sort-Object Caption
        
            $htmlbody += $process | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect Printers information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Printers Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting printers information"
        
        try
        {
            $printer = Get-WmiObject Win32_Printer -ComputerName $ComputerName -ErrorAction STOP | Select-Object Caption,Description,DriverName,JobCountSinceLastReset,Network,PortName,PrintProcessor,Published,StartTime,TimeOfLastReset,ServerName,ShareName | Sort-Object Caption
        
            $htmlbody += $printer | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect Printer Drivers information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Printer Drivers Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting Printer Drivers information"
        
        try
        {
            $printerdriver = Get-WmiObject Win32_PrinterDriver -ComputerName $ComputerName -ErrorAction STOP | Select-Object Name,ConfigFile,Description,DriverPath,MonitorName,Version | Sort-Object Name
        
            $htmlbody += $printerdriver | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect Printer Ports information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Printer Ports Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting Printer Ports information"
        
        try
        {
            $printerport = Get-WmiObject Win32_TcpIpPrinterPort -ComputerName $ComputerName -ErrorAction STOP | Select-Object Caption,Description,HostAddress,Name,PortNumber,Protocol,Queue,Status,SystemName,Type | Sort-Object Name
        
            $htmlbody += $printerport | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect Network Shares information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Network Shares Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting network shares information"
        
        try
        {
            $share = Get-WmiObject Win32_Share -ComputerName $ComputerName -ErrorAction STOP | Select-Object Caption,Description,Name,Path,Type | Sort-Object Caption
        
            $htmlbody += $share | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }
              

        #---------------------------------------------------------------------
        # Collect Scheduled Tasks information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Scheduled Tasks Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting Scheduled Tasks information"
        
        try
        {
            $scheduledtask = Get-WmiObject Win32_ScheduledJob -ComputerName $ComputerName -ErrorAction STOP | Select-Object Caption,Command,DaysOfMonth,DaysOfWeek,Description,ElapsedTime,InstallDate,Name,Owner,RunRepeatedly,StartTime,Status,TimeSubmitted,Untiltime | Sort-Object Caption
        
            $htmlbody += $scheduledtask | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect IIS information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>IIS Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting Internet Information Server (IIS) information 2"
        
        try
        {
           $iisversion =Invoke-Command -ComputerName $ComputerName -ScriptBlock {  $(get-itemproperty HKLM:\SOFTWARE\Microsoft\InetStp\).setupstring}
        }
        catch
        {
            Write-Warning $_.Exception.Message
        }

        If ($iisversion -like '*IIS 6*')
        {
            Write-Host This server uses IIS6
        $IISWebInfo1 = Get-WmiObject -class "IIsWebInfo" -namespace "root\microsoftiisv2" -ComputerName $ComputerName -ErrorAction STOP | Select-Object Name, MajorIIsVersionNumber, MinorIIsVersionNumber | Sort-Object Name

        $IISWebInfo2 = invoke-command -computername  $ComputerNAme { Import-Module WebAdministration; Get-ChildItem –Path IIS:\Sites} -ErrorAction STOP | Select-Object name, id, serverAutoStart, state, applicationPool, enabledProtocols, physicalPath | Sort-Object name

        $IISWebInfo4 = Get-WmiObject -class "IIsWebVirtualDirSetting" -namespace "root\microsoftiisv2" -ComputerName $ComputerName -ErrorAction STOP | Select-Object Name, ServerComment, AppFriendlyName, AppPoolId, AspScriptLanguage, Bindings, Caption, DefaultDoc, Description, Path, ScriptMaps | Sort-Object Name

        $IISWebInfo6 = Get-WmiObject -class "IIsWebServiceSetting" -namespace "root\microsoftiisv2" -ComputerName $ComputerName -ErrorAction STOP | Select-Object Name, ServerComment, AppFriendlyName, AppPoolId, AspScriptLanguage, Bindings, Caption, DefaultDoc, Description, Path, ScriptMaps, ServerSize, ServerBindings | Sort-Object Name

        $IISWebInfo7 = invoke-command -computername  $ComputerNAme { Import-Module WebAdministration; Get-ChildItem –Path IIS:\AppPools} -ErrorAction STOP | Select-Object name, autoStart, managedRuntimeVersion, startMode, state

        $IISWebInfo8 = Get-WmiObject -class "IIsApplicationPoolsSetting" -namespace "root\microsoftiisv2" -ComputerName $ComputerName -ErrorAction STOP | Select-Object Name, Caption, Description, AppPoolCommand, AppPoolIdentityType, AppPoolState | Sort-Object Name

        $IISWebInfo9 = Get-WmiObject -class "IIsFilterSetting" -namespace "root\microsoftiisv2" -ComputerName $ComputerName -ErrorAction STOP | Select-Object Name, Caption, Description, FilterDescription, FilterPath, FilterEnabled | Sort-Object Name
        
        $IISWebInfo10 = Get-WmiObject -class "IIsFtpServerSetting" -namespace "root\microsoftiisv2" -ComputerName $ComputerName -ErrorAction STOP | Select-Object Name, Caption, Description, ServerComment | Sort-Object Name

        $subhead = "<h3>IIS Information - General</h3>"
        $htmlbody += $subhead 
        
        $htmlbody += $IISWebInfo1 | ConvertTo-Html -Fragment
        $htmlbody += $spacer

        $subhead = "<h3>IIS Information - Web Sites</h3>"
        $htmlbody += $subhead 

        $htmlbody += $IISWebInfo | ConvertTo-Html -Fragment
        $htmlbody += $spacer

        $subhead = "<h3>IIS Information - Virtual Directories</h3>"
        $htmlbody += $subhead 

        $htmlbody += $IISWebInfo4 | ConvertTo-Html -Fragment
        $htmlbody += $spacer

        $subhead = "<h3>IIS Information - Service</h3>"
        $htmlbody += $subhead

        $htmlbody += $IISWebInfo6 | ConvertTo-Html -Fragment
        $htmlbody += $spacer

        $subhead = "<h3>IIS Information - Application Pools</h3>"
        $htmlbody += $subhead 

        $htmlbody += $IISWebInfo7 | ConvertTo-Html -Fragment
        $htmlbody += $spacer 

        $htmlbody += $IISWebInfo8 | ConvertTo-Html -Fragment
        $htmlbody += $spacer 

        $subhead = "<h3>IIS Information - Filters</h3>"
        $htmlbody += $subhead 

        $htmlbody += $IISWebInfo9 | ConvertTo-Html -Fragment
        $htmlbody += $spacer 

        $subhead = "<h3>IIS Information - FTP</h3>"
        $htmlbody += $subhead 

        $htmlbody += $IISWebInfo10 | ConvertTo-Html -Fragment
        $htmlbody += $spacer 

        }
        else
        {

        try
        {
        
            $IISWebInfo1 = invoke-command -computername  $ComputerNAme { Import-Module WebAdministration; Get-ItemProperty IIS:\ -Name applicationPoolDefaults} -ErrorAction STOP | Select-Object * | Sort-Object name

            $IISWebInfo2 = invoke-command -computername  $ComputerNAme { Import-Module WebAdministration; Get-ChildItem –Path IIS:\Sites} -ErrorAction STOP | Select-Object name, id, serverAutoStart, state, applicationPool, enabledProtocols, physicalPath, Bindings | Sort-Object name
                        
            $IISWebInfo3 = invoke-command -computername  $ComputerNAme { Import-Module WebAdministration; Get-ChildItem –Path IIS:\AppPools} -ErrorAction STOP | Select-Object name, state, CLRConfigFile, managedRuntimeVersion, managedPipelineMode, startMode | Sort-Object name
            
            $IISWebInfo4 = invoke-command -computername  $ComputerNAme { Import-Module WebAdministration; Get-WebConfiguration system.webServer/security/authentication/* 'IIS:\sites\' -Recurse} -ErrorAction STOP | Select-Object *

            $IISWebInfo5 = invoke-command -computername  $ComputerNAme { Import-Module WebAdministration; Get-ChildItem –Path IIS:\SSLBindings} -ErrorAction STOP | Select-Object *
                        
            $subhead = "<h3>IIS Information - Applications Pools defaults</h3>"
            $htmlbody += $subhead 
        
            $htmlbody += $IISWebInfo1 | ConvertTo-Html -Fragment
            $htmlbody += $spacer

            $subhead = "<h3>IIS Information - Web Sites</h3>"
            $htmlbody += $subhead 

            $htmlbody += $IISWebInfo2 | ConvertTo-Html -Fragment
            $htmlbody += $spacer

            $subhead = "<h3>IIS Information - Application Pools</h3>"
            $htmlbody += $subhead 

            $htmlbody += $IISWebInfo3 | ConvertTo-Html -Fragment
            $htmlbody += $spacer

            $subhead = "<h3>IIS Information - Security</h3>"
            $htmlbody += $subhead

            $htmlbody += $IISWebInfo4 | ConvertTo-Html -Fragment
            $htmlbody += $spacer

            $subhead = "<h3>IIS Information - Applications</h3>"
            $htmlbody += $subhead 

            $htmlbody += $IISWebInfo5 | ConvertTo-Html -Fragment
            $htmlbody += $spacer         
        }
        
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }
        }


        #---------------------------------------------------------------------
        # Collect TCP Connections information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>TCP Connections Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting TCP Connections information"
        
        try
        {
            $TCPConnectionInfo = Get-NetTCPConnection -ErrorAction STOP | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess, CreationTime | Sort-Object LocalAddress
        
            $htmlbody += $TCPConnectionInfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect ODBC Drivers information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>ODBC Drivers Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting ODBC Drivers information"

        try
        {

	    $Base = New-Object PSObject;
	    $Base | Add-Member Noteproperty Name -Value $Null;
	    $Base | Add-Member Noteproperty Driver -Value $Null;
	    $Base | Add-Member Noteproperty DriverODBCVer -Value $Null;
	    $Results =  New-Object System.Collections.Generic.List[System.Object];
        
		$Registry = $Null;

		try
        {
        
            $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$ComputerName);
        
        }
		
        catch
        {
        
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        
        }
		
		if ($Registry)
            {
			
            $ODBCKeys = $Null;
			$SubKey = $Null;
			$ODBCKeys = $Registry.OpenSubKey("Software\odbc\odbcinst.ini\",$False);
			$ODBCKeys.GetSubKeyNames()|%{
			$ODBCSubKey = $ODBCKeys.OpenSubKey($_,$False);
            
			$Entry = $Base | Select-Object *
            $Entry.Name = $ODBCSubKey
			$Entry.Driver = $ODBCSubKey.GetValue("Driver");
            $Entry.DriverODBCVer = $ODBCSubKey.GetValue("DriverODBCVer");

            [Void]$Results.Add($Entry);

		    }

            }
		
        $htmlbody += $results | ConvertTo-Html -Fragment
        $htmlbody += $spacer 

        }

        catch
        {

            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer

        }


        #---------------------------------------------------------------------
        # Collect ODBC Sources information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>ODBC Sources Information</h3>"
        $htmlbody += $subhead
 
        Write-Verbose "Collecting ODBC Data sources information"

        try
        {

	    $Base = New-Object PSObject;
	    $Base | Add-Member Noteproperty Name -Value $Null;
	    $Base | Add-Member Noteproperty Driver -Value $Null;
	    $Base | Add-Member Noteproperty Server -Value $Null;
	    $Results =  New-Object System.Collections.Generic.List[System.Object];
        
		$Registry = $Null;

		try
        {
        
            $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$ComputerName);
        
        }
		
        catch
        {
        
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer

        
        }
		
		if ($Registry)
            {
			
            $ODBCSourcesKeys = $Null;
			$ODBCSourcesSubKey = $Null;
			$ODBCSourcesKeys = $Registry.OpenSubKey("Software\odbc\odbc.ini\",$False);
			$ODBCSourcesKeys.GetSubKeyNames()|%{
			$ODBCSourcesSubKey = $ODBCSourcesKeys.OpenSubKey($_,$False);
            
			$Entry = $Base | Select-Object *
            $Entry.Name = $ODBCSourcesSubKey
			$Entry.Driver = $ODBCSourcesSubKey.GetValue("Driver");
            $Entry.Server = $ODBCSourcesSubKey.GetValue("Server");

            [Void]$Results.Add($Entry);

		        }

            $ODBC6432SourcesKeys = $Null;
			$ODBC6432SourcesSubKey = $Null;
			$ODBC6432SourcesKeys = $Registry.OpenSubKey("Software\WOW6432Node\odbc\odbc.ini\",$False);
			$ODBC6432SourcesKeys.GetSubKeyNames()|%{
			$ODBC6432SourcesSubKey = $ODBC6432SourcesKeys.OpenSubKey($_,$False);
            
			$Entry = $Base | Select-Object *
            $Entry.Name = $ODBC6432SourcesSubKey
			$Entry.Driver = $ODBC6432SourcesSubKey.GetValue("Driver");
            $Entry.Server = $ODBC6432SourcesSubKey.GetValue("Server");

            [Void]$Results.Add($Entry);

		        }


            }
		

        $htmlbody += $results | ConvertTo-Html -Fragment
        $htmlbody += $spacer 

        }

        catch
        {

            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer

        }

        #---------------------------------------------------------------------
        # Generate the HTML report and output to file
        #---------------------------------------------------------------------
	
        Write-Verbose "Producing HTML report"
    
        $reportime = Get-Date

        #Common HTML head and styles
	    $htmlhead="<html>
				    <style>
				    BODY{font-family: Arial; font-size: 8pt;}
				    H1{font-size: 20px;}
				    H2{font-size: 18px;}
				    H3{font-size: 16px;}
				    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				    TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
				    TD{border: 1px solid black; padding: 5px; }
				    td.pass{background: #7FFF00;}
				    td.warn{background: #FFE600;}
				    td.fail{background: #FF0000; color: #ffffff;}
				    td.info{background: #85D4FF;}
				    </style>
				    <body>
				    <h1 align=""center"">Server Info: $ComputerName</h1>
				    <h3 align=""center"">Generated: $reportime</h3>"

        $htmltail = "</body>
			    </html>"

        $htmlreport = $htmlhead + $htmlbody + $htmltail

        $htmlreport | Out-File $htmlfile -Encoding Utf8
    }
    }
}

End
{
    #Wrap it up
    Write-Verbose "=====> Finished <====="
}
