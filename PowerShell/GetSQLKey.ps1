Function GetSqlServerKey
{
    [CmdletBinding(DefaultParameterSetName = "Default")]
    Param (
        [parameter(Position = 0)]
        [Alias("ServerInstance","SqlInstance")]
        [string[]]$SqlServers,
        # Central Management Server

        [string]$SqlCms,
        # File with one server per line

        [string]$ServersFromFile,
        [System.Management.Automation.PSCredential]$SqlCredential
    )
    
    BEGIN
    {
        
        Function Unlock-SqlServerKey
        {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory = $true)]
                [byte[]]$data,
                [int]$version
            )
            try
            {
                if ($version -ge 11) { $binArray = ($data)[0..66] }
                else { $binArray = ($data)[52..66] }
                $charsArray = "B", "C", "D", "F", "G", "H", "J", "K", "M", "P", "Q", "R", "T", "V", "W", "X", "Y", "2", "3", "4", "6", "7", "8", "9"
                for ($i = 24; $i -ge 0; $i--)
                {
                    $k = 0
                    for ($j = 14; $j -ge 0; $j--)
                    {
                        $k = $k * 256 -bxor $binArray[$j]
                        $binArray[$j] = [math]::truncate($k / 24)
                        $k = $k % 24
                    }
                    $productKey = $charsArray[$k] + $productKey
                    if (($i % 5 -eq 0) -and ($i -ne 0))
                    {
                        $productKey = "-" + $productKey
                    }
                }
            }
            catch { $productkey = "Cannot decode product key." }
            return $productKey
        }
    }
    
    PROCESS
    {
        
        if ($SqlCms)
        {
            if ([Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Management.RegisteredServers") -eq $null)
            { throw "Can't load CMS assemblies. You must have SQL Server Management Studio installed to use the -SqlCms switch." }
            
            Write-Output "Gathering SQL Servers names from Central Management Server"
            $server = Connect-SqlServer -SqlServer $SqlCms -SqlCredential $SqlCredential
            $sqlconnection = $server.ConnectionContext.SqlConnectionObject
            
            try { $cmstore = New-Object Microsoft.SqlServer.Management.RegisteredServers.RegisteredServersStore($sqlconnection) }
            catch { throw "Cannot access Central Management Server" }
            $dbstore = $cmstore.DatabaseEngineServerGroup
            $SqlServers = $dbstore.GetDescendantRegisteredServers().servername
            # Add the CM server itself, which can't be stored in the CM server.
            $servers += $SqlCms
            $basenames = @()
            foreach ($server in $SqlServers) { $basenames += $server.Split("\")[0] }
            $SqlServers = $basenames | Get-Unique
        }
        
        If ($ServersFromFile)
        {
            if ((Test-Path $ServersFromFile) -eq $false) { throw "Could not find file: $ServersFromFile" }
            $SqlServers = Get-Content $ServersFromFile
        }
        
        if ([string]::IsNullOrEmpty($SqlServers)) { $SqlServers = $env:computername }
        
        $basepath = "SOFTWARE\Microsoft\Microsoft SQL Server"
        # Loop through each server
        $objectCollection = @()
        foreach ($servername in $SqlServers)
        {
            $servername = $servername.Split("\")[0]
            
            if ($servername -eq "." -or $servername -eq "localhost" -or $servername -eq $env:computername)
            {
                $localmachine = [Microsoft.Win32.RegistryHive]::LocalMachine
                $defaultview = [Microsoft.Win32.RegistryView]::Registry64
                $reg = [Microsoft.Win32.RegistryKey]::OpenBaseKey($localmachine, $defaultview)
            }
            else
            {
                # Get IP for remote registry access. It's the most reliable.
                try { $ipaddr = ([System.Net.Dns]::GetHostAddresses($servername)).IPAddressToString }
                catch { Write-Warning "Can't resolve $servername. Skipping."; continue }
                
                try
                {
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $ipaddr)
                }
                catch { Write-Warning "Can't access registry for $servername. Is the Remote Registry service started?"; continue }
            }
            
            $instances = $reg.OpenSubKey("$basepath\Instance Names\SQL", $false)
            if ($instances -eq $null) { Write-Warning "No instances found on $servername. Skipping."; continue }
            # Get Product Keys for all instances on the server.
            foreach ($instance in $instances.GetValueNames())
            {
                if ($instance -eq "MSSQLSERVER") { $sqlserver = $servername }
                else { $sqlserver = "$servername\$instance" }
                
                $subkeys = $reg.OpenSubKey("$basepath", $false)
                $instancekey = $subkeys.GetSubKeynames() | Where-Object { $_ -like "*.$instance" }
                if ($instancekey -eq $null) { $instancekey = $instance } # SQL 2k5
                
                # Cluster instance hostnames are required for SMO connection
                $cluster = $reg.OpenSubKey("$basepath\$instancekey\Cluster", $false)
                if ($cluster -ne $null)
                {
                    $clustername = $cluster.GetValue("ClusterName")
                    if ($instance -eq "MSSQLSERVER") { $sqlserver = $clustername }
                    else { $sqlserver = "$clustername\$instance" }
                }
                
                Write-Verbose "Attempting to connect to $sqlserver"
                try { $server = Connect-SqlServer -SqlServer $SqlServer -SqlCredential $SqlCredential }
                catch { Write-Warning "Can't connect to $sqlserver or access denied. Skipping."; continue }
                
                $servicePack = $server.ProductLevel
                Write-Debug "$servername $instance version is $($server.VersionMajor)"
                switch ($server.VersionMajor)
                {                    
                    9 {
                        $sqlversion = "SQL Server 2005 $servicePack"
                        $findkeys = $reg.OpenSubKey("$basepath\90\ProductID", $false)
                        foreach ($findkey in $findkeys.GetValueNames())
                        {
                            if ($findkey -like "DigitalProductID*") { $key = "$basepath\90\ProductID\$findkey" }
                        }
                    }
                    10 {
                        $sqlversion = "SQL Server 2008 $servicePack"
                        $key = "$basepath\MSSQL10"
                        if ($server.VersionMinor -eq 50) { $key += "_50"; $sqlversion = "SQL Server 2008 R2 $servicePack" }
                        $key += ".$instance\Setup\DigitalProductID"
                    }
                    11 { $key = "$basepath\110\Tools\Setup\DigitalProductID"; $sqlversion = "SQL Server 2012 $servicePack" }
                    12 { $key = "$basepath\120\Tools\Setup\DigitalProductID"; $sqlversion = "SQL Server 2014 $servicePack" }
                    13 { $key = "$basepath\130\Tools\Setup\DigitalProductID"; $sqlversion = "SQL Server 2016 $servicePack" }
                    default { Write-Warning "SQL version not currently supported."; continue }
                }
                if ($server.Edition -notlike "*Express*")
                {
                    try
                    {
                        $subkey = Split-Path $key; $binaryvalue = Split-Path $key -leaf
                        $binarykey = $($reg.OpenSubKey($subkey)).GetValue($binaryvalue)
                    }
                    catch { $sqlkey = "Could not connect." }
                    try { $sqlkey = Unlock-SqlServerKey $binarykey $server.VersionMajor }
                    catch { }
                }
                else { $sqlkey = "SQL Server Express Edition" }
                $server.ConnectionContext.Disconnect()
                
                ##$object = New-Object PSObject -Property @{
                ##    "SQL Instance" = $sqlserver
                ##    "SQL Version" = $sqlversion
                ##    "SQL Edition" = $server.Edition
                ##    "Product Key" = $sqlkey
                ##}
                
                $win32os = Get-WmiObject Win32_OperatingSystem -computer $target
                $ipV4 = Test-Connection -ComputerName ($env:computerName) -Count 1  | Select -ExpandProperty IPV4Address
               	$ADDomain = Get-ADDomain
		        $obj = New-Object Object
		        $obj | Add-Member Noteproperty Computer -value "$($env:computerName . $ADDomain)"
		        $obj | Add-Member Noteproperty "OS Architecture" -value $win32os.OSArchitecture
		        $obj | Add-Member Noteproperty Caption -value $win32os.Caption
		        ## $obj | Add-Member Noteproperty CSDVersion -value $win32os.CSDVersion
		        $obj | Add-Member Noteproperty OSArch -value $win32os.OSArchitecture
		        $obj | Add-Member Noteproperty BuildNumber -value $win32os.BuildNumber
		        $obj | Add-Member Noteproperty RegisteredTo -value $win32os.RegisteredUser
		        $obj | Add-Member Noteproperty "ProductID (Windows)" -value $win32os.SerialNumber
		        $obj | Add-Member Noteproperty "SQL IP" -value $ipV4
		        $obj | Add-Member Noteproperty "SQL Instance" -value $sqlserver
		        $obj | Add-Member Noteproperty "SQL Version" -value $sqlversion
		        $obj | Add-Member Noteproperty "SQL Edition" -value $server.Edition
		        $obj | Add-Member Noteproperty "ProductKey (SQL Server)" -value $sqlkey
		        $obj
                
                
                ## $objectCollection += $object
            }
            $reg.Close()
        }
        ## $objectCollection | Select "SQL Instance", "SQL Version", "SQL Edition", "Product Key"
    }
    
    END
    {
        # Write-Output "Script completed"
    }
}

GetSqlServerKey