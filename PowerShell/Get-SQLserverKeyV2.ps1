Function Get-SQLserverKey {
    ## function to retrieve the Windows Product Key from any PC as well as the SQL Server Product Key
    Param ($targets = ".")
    $hklm = 2147483650
    
    $basepath = "SOFTWARE\Microsoft\Microsoft SQL Server"
    $regPath2005 = "$basepath\90\ProductID"
    $regPath2005version = "$basepath\90\Tools\ClientSetup\CurrentVersion"
    $regPath2008 = "$basepath\100\Tools\Setup"
    $regPath2012 = "$basepath\110\Tools\Setup"
    $regPath2014 = "$basepath\120\Tools\Setup"
    $regPath2016 = "$basepath\130\Tools\Setup"
    $regValue2005 = "DigitalProductId"
    $regValue2008 = "DigitalProductId"
    $regValue2012 = "DigitalProductId"
    $regValue2014 = "DigitalProductId"
    $regValue2016 = "DigitalProductId"
    $regValueVersion2005 = "CurrentVersion"
    $regValueVersion2008 = "Version"
    $regValueVersion2012 = "Version"
    $regValueVersion2014 = "Version"
    $regValueVersion2016 = "Version"
    $SQLInstanceKey = "$basepath\Instance Names"
    $SQLinstalled = $false
    Foreach ($target in $targets) {
    	$wmi = [WMIClass]"\\$target\root\default:stdRegProv"
    	#There are a few types of SQL install, and the information is in one of these
        ForEach ($Type in ("SQL","RS","OLAP"))
        {
    		$TypeKey = "$SQLInstanceKey\$Type"
            $Instance = ($wmi.EnumValues($hklm,$TypeKey)).sNames | Select -First 1
            
            If ($Instance)
            {
            	$InstanceName = $wmi.GetStringValue($hklm,$TypeKey,$Instance).sValue
            	Break
            }
		}	
    
        $productKey = $null
        $win32os = $null
        $sqlversion = $null
        $sqledition = $null
        $sqlinstance = $InstanceName
        
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
        	$regPath2005 = "$basepath\$InstanceName\Setup"
        	$regPath2005version = "$basepath\$InstanceName\ClientSetup"
        	$regValueVersion2005 = "Version"
            $data2005 = $wmi.GetBinaryValue($hklm,$regPath2005,$regValue2005)
        }
        $data2008 = $wmi.GetBinaryValue($hklm,$regPath2008,$regValue2008)
        $data2012 = $wmi.GetBinaryValue($hklm,$regPath2012,$regValue2012)
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
        $win32os = Get-WmiObject Win32_OperatingSystem -computer $target
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
    }
}

Get-SQLserverKey