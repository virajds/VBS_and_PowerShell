Function Get-SQLserverKey {
    ## function to retrieve the Windows Product Key from any PC as well as the SQL Server Product Key
    Param ($targets = ".")
    $hklm = 2147483650
    $regPath2005 = "SOFTWARE\Microsoft\Microsoft SQL Server\90\ProductID"
    $regPath2005version = "SOFTWARE\Microsoft\Microsoft SQL Server\90\Tools\ClientSetup\CurrentVersion"
    $regPath2008 = "SOFTWARE\Microsoft\Microsoft SQL Server\100\Tools\Setup"
    $regPath2012 = "SOFTWARE\Microsoft\Microsoft SQL Server\110\Tools\Setup"
    $regValue2005 = "DigitalProductId77591"
    $regValue2008 = "DigitalProductId"
    $regValue2012 = "DigitalProductId"
    $regValueVersion2005 = "CurrentVersion"
    $regValueVersion2008 = "Version"
    $regValueVersion2012 = "Version"
    $SQLinstalled = $false
    Foreach ($target in $targets) {
        $productKey = $null
        $win32os = $null
        $sqlversion = $null
        $sqledition = $null
        $sqlinstance = [System.Data.Sql.SqlDataSourceEnumerator]::Instance.GetDataSources()|?{$_.ServerName -eq $env:COMPUTERNAME}
        $wmi = [WMIClass]"\\$target\root\default:stdRegProv"
        $data2005 = $wmi.GetBinaryValue($hklm,$regPath2005,$regValue2005)
        if ($data2005.uValue.length -lt 1) {
        	$regValue2005 = "DigitalProductId77671"
        	$data2005 = $wmi.GetBinaryValue($hklm,$regPath2005,$regValue2005)
        }
        $data2008 = $wmi.GetBinaryValue($hklm,$regPath2008,$regValue2008)
        $data2012 = $wmi.GetBinaryValue($hklm,$regPath2012,$regValue2012)
        $productKey = ""
        if ($data2005.uValue.length -ge 1) { 
            $binArray2005 = ($data2005.uValue)[52..66]
            $binArray = $binArray2005
            $sqlversion = $wmi.GetStringValue($hklm,$regPath2005version,$regValueVersion2005).sValue
            $sqledition = $wmi.GetStringValue($hklm,$regPath2005version,$regValueVersion2005).Edition
            $SQLinstalled = $true
        }
        if ($data2008.uValue.length -ge 1) { 
            $binArray2008 = ($data2008.uValue)[52..66]
            $binArray = $binArray2008
            $sqlversion = $wmi.GetStringValue($hklm,$regPath2008,$regValueVersion2008).sValue
            $sqledition = $wmi.GetStringValue($hklm,$regPath2008,$regValueVersion2008).Edition
            $SQLinstalled = $true
        }
        if ($data2012.uValue.length -ge 1) { 
            $binArray2012 = ($data2012.uValue)
            $binArray = $binArray2012
            $sqlversion = $wmi.GetStringValue($hklm,$regPath2012,$regValueVersion2012).sValue
			$sqledition = $wmi.GetStringValue($hklm,$regPath2012,$regValueVersion2012).Edition
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
        $obj | Add-Member Noteproperty "SQLedition" -value $sqlversion
        $obj | Add-Member Noteproperty "SQLinstance" -value $sqlinstance
        $obj | Add-Member Noteproperty "ProductKey" -value $productkey
        $obj
    }
}

Get-SQLserverKey