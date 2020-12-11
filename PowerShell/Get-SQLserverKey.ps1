Function Get-SQLserverKey {
	## function to retrieve the Windows Product Key from any PC as well as the SQL Server Product Key
	
	$basepath = "SOFTWARE\Microsoft\Microsoft SQL Server"
    $wmi = [WMIClass]"\\.\root\default:stdRegProv"
	
	$servername = $env:COMPUTERNAME
    $hklm = 2147483650
    $VersionKey  = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $SQLInstanceKey = "$basepath\Instance Names"
    
    $regPath2005 = "SOFTWARE\Microsoft\Microsoft SQL Server\90\ProductID"
    $regPath2005version = "SOFTWARE\Microsoft\Microsoft SQL Server\90\Tools\ClientSetup\CurrentVersion"
    $regPath2008 = "SOFTWARE\Microsoft\Microsoft SQL Server\100\Tools\Setup"
    $regPath2012 = "SOFTWARE\Microsoft\Microsoft SQL Server\110\Tools\Setup"
    $regPath2014 = "SOFTWARE\Microsoft\Microsoft SQL Server\120\Tools\Setup"
    $regPath2016 = "SOFTWARE\Microsoft\Microsoft SQL Server\130\Tools\Setup"
    $regValue2005 = "DigitalProductId77591"
    $regValue2008 = "DigitalProductId"
    $regValue2012 = "DigitalProductId"
    $regValue2014 = "DigitalProductId"
    $regValue2016 = "DigitalProductId"
    $regValueVersion2005 = "CurrentVersion"
    $regValueVersion2008 = "CurrentVersion"
    $regValueVersion2008R2 = "Version"
    $regValueVersion2012 = "Version"
    $regValueVersion2014 = "Version"
    $regValueVersion2016 = "Version"
    $SQLinstalled = $false
		
	#Check if SQL is installed
    $SQLInstalls = @($wmi.EnumKey($hklm,$VersionKey) | Select -ExpandProperty SNames | Where { $_ -like "*SQL*" })
    
    If (!$SQLInstalls)
    {
    	Write-Output "No SQL instalation found"
    }
    
    If ($SQLInstalls)
    {
		#Good, it is.  Get Name, determine if this is a full SQL install (or a just a SSMS install)
        ForEach ($Install in $SQLInstalls)
        {
            $Name = ""
            $UninstallString = "$VersionKey\$Install"
            If ($wmi.GetStringValue($hklm,$UninstallString,"UninstallString").sValue)
            {
                $Name = $wmi.GetStringValue($hklm,$UninstallString,"DisplayName").sValue
                Break
            }
        }
        
        If ($Name)
        {
            #There are a few types of SQL install, and the information is in one of these
            ForEach ($Type in ("SQL","RS","OLAP"))
            {
        		$TypeKey = "$SQLInstanceKey\$Type"
                $Instance = ($wmi.EnumValues($hklm,$TypeKey)).sNames | Select -First 1
                
                If ($Instance)
                {
                	$InstanceName = $wmi.GetStringValue($hklm,$TypeKey,$Instance).sValue
                	
					if ($InstanceName -eq "MSSQLSERVER") { $sqlserver = $servername }
			        else { $sqlserver = "$servername\$InstanceName" }
			        
			        #$subkeys = $wmi.EnumKey($hklm,"$basepath")
			        #$instancekey = $subkeys.EnumKey() | Where-Object { $_ -like "*.$InstanceName" }
			        #if ($instancekey -eq $null) { $instancekey = $InstanceName } # SQL 2k5
			        
			        # Cluster instance hostnames are required for SMO connection
			        #$cluster = $wmi.EnumKey($hklm,"$basepath\$instancekey\Cluster")
			        #if ($cluster -ne $null)
			        #{
			        #    $clustername = $cluster.GetStringValue("ClusterName")
			        #    if ($InstanceName -eq "MSSQLSERVER") { $sqlserver = $clustername }
			        #    else { $sqlserver = "$clustername\$InstanceName" }
			        #}
			        
			        $productKey = $null
			        $win32os = $null
			        $sqlversion = $null
			        $sqledition = $null
			        $sqlinstance = $sqlserver
			        $data2005 = $wmi.GetBinaryValue($hklm,$regPath2005,$regValue2005)
			        if ($data2005.uValue.length -lt 1) {
			        	$findkeys = $wmi.EnumKey($hklm,"$basepath\90\ProductID")
			            foreach ($findkey in ($findkeys.sNames))
			            {
			                if ($findkey -like "DigitalProductID*") { $regValue2005 = $findkey }
			                $data2005 = $wmi.GetBinaryValue($hklm,$regPath2005,$regValue2005)
			            }
			        }
			        if ($data2005.uValue.length -lt 1) {
			        	$regPath2005 = "$basepath\MSSQL10_50.MSSQLSERVER2008\Setup\DigitalProductID"
			            $data2005 = $wmi.GetBinaryValue($hklm,$regPath2005,$regValue2005)
			        }
			        if ($data2005.uValue.length -lt 1) {
			        	$regPath2005 = "$basepath\MSSQL10_50.SQLEXPRESS2008R2\Setup\DigitalProductID"
			            $data2005 = $wmi.GetBinaryValue($hklm,$regPath2005,$regValue2005)
			        }
			        $data2008 = $wmi.GetBinaryValue($hklm,$regPath2008,$regValue2008)
			        if ($data2008.uValue.length -lt 1) {
			        	$regPath2008 = "$basepath\MSSQL10_50.MSSQLSERVER2008\Setup\DigitalProductID"
			            $data2008 = $wmi.GetBinaryValue($hklm,$regPath2008,$regValue2008)
			        }
			        if ($data2008.uValue.length -lt 1) {
			        	$regPath2008 = "$basepath\MSSQL10_50.SQLEXPRESS2008R\Setup\DigitalProductID"
			            $data2008 = $wmi.GetBinaryValue($hklm,$regPath2008,$regValue2008)
			        }
			        $data2012 = $wmi.GetBinaryValue($hklm,$regPath2012,$regValue2012)
			        $data2014 = $wmi.GetBinaryValue($hklm,$regPath2014,$regValue2016)
			        $data2016 = $wmi.GetBinaryValue($hklm,$regPath2014,$regValue2016)
			        $productKey = ""
			        if ($data2005.uValue.length -ge 1) { 
			            $binArray2005 = ($data2005.uValue)[52..66]
			            $binArray = $binArray2005
			            $sqlversion = $wmi.GetStringValue($hklm,$regPath2005version,$regValueVersion2005).sValue
			            $sqledition = $wmi.GetStringValue($hklm,$regPath2005version,"Edition").sValue
			            $SQLinstalled = $true
			        }
			        if ($data2008.uValue.length -ge 1) { 
			            $binArray2008 = ($data2008.uValue)[52..66]
			            $binArray = $binArray2008
			            $sqlversion = $wmi.GetStringValue($hklm,$regPath2008,$regValueVersion2008).sValue
			            $sqledition = $wmi.GetStringValue($hklm,$regPath2008,"Edition").sValue
			            $SQLinstalled = $true
			        }
			        if ($data2012.uValue.length -ge 1) { 
			            $binArray2012 = ($data2012.uValue)[0..66]
			            $binArray = $binArray2012
			            $sqlversion = $wmi.GetStringValue($hklm,$regPath2012,$regValueVersion2012).sValue
			            $sqledition = $wmi.GetStringValue($hklm,$regPath2012,"Edition").sValue
			            $SQLinstalled = $true
					}
					if ($data2014.uValue.length -ge 1) { 
			            $binArray2014 = ($data2014.uValue)[0..66]
			            $binArray = $binArray2014
			            $sqlversion = $wmi.GetStringValue($hklm,$regPath2014,$regValueVersion2014).sValue
			            $sqledition = $wmi.GetStringValue($hklm,$regPath2014,"Edition").sValue
			            $SQLinstalled = $true
					}
					if ($data2016.uValue.length -ge 1) { 
			            $binArray2016 = ($data2016.uValue)[0..66]
			            $binArray = $binArray2016
			            $sqlversion = $wmi.GetStringValue($hklm,$regPath2016,$regValueVersion2016).sValue
			            $sqledition = $wmi.GetStringValue($hklm,$regPath2016,"Edition").sValue
			            $SQLinstalled = $true
					}
			
			        if ($SQLinstalled) {
			            ## decrypt base24 encoded binary data
			            $charsArray = "B","C","D","F","G","H","J","K","M","P","Q","R","T","V","W","X","Y","2","3","4","6","7","8","9"
			            For ($i = 24; $i -ge 0; $i--) {
			                $k = 0
			                For ($j = 14; $j -ge 0; $j--) {
			                    $k = $k * 256 -bxor $binArray[$j]
			                    $binArray[$j] = [math]::truncate($k / 24)
			                    $k = $k % 24
			                }
			                $productKey = $charsArray[$k] + $productKey
			                If (($i % 5 -eq 0) -and ($i -ne 0)) {
			                    $productKey = "-" + $productKey
			                }
			            }
			        } else {
			            $productKey = "no SQL Server found"
			        }
			        $win32os = Get-WmiObject Win32_OperatingSystem -computer .
			        ## $ipV4 = Test-Connection -ComputerName ($env:computerName) -Count 1  | Select -ExpandProperty IPV4Address
			        
			        $obj = New-Object Object
			        $obj | Add-Member Noteproperty Computer -value "$env:COMPUTERNAME.$env:USERDNSDOMAIN"
			        $obj | Add-Member Noteproperty OSCaption -value $win32os.Caption
			        ## $obj | Add-Member Noteproperty CSDVersion -value $win32os.CSDVersion
			        $obj | Add-Member Noteproperty OSArch -value $win32os.OSArchitecture
			        ##$obj | Add-Member Noteproperty BuildNumber -value $win32os.BuildNumber
			        ##$obj | Add-Member Noteproperty RegisteredTo -value $win32os.RegisteredUser
			        ##$obj | Add-Member Noteproperty "ProductID (Windows)" -value $win32os.SerialNumber
			        $obj | Add-Member Noteproperty "SQLVer" -value $sqlversion
			        $obj | Add-Member Noteproperty "SQLedition" -value $sqledition
			        $obj | Add-Member Noteproperty "SQLinstance" -value $sqlinstance
			        $obj | Add-Member Noteproperty "ProductKey" -value $productkey
			        $obj
			        Break
				}
			}
		}
    }
}

Get-SQLserverKey