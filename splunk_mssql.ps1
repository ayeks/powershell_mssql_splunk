<#
        .SYNOPSIS
            Generates Splunk Config for MSSQL Servers

        .DESCRIPTION
            Retrieves SQL server information and writes Splunk Config in predefined Dirs.
            Works with multiple instances on one server or cluster.
            Adds Errorlogpath and Agentlogpath to inputs.conf

        .PARAMETER DirApp
            Directory to Splunk MSSQL App. Content will be deletet. Default path: "C:\Program Files\SplunkUniversalForwarder\etc\apps\mssql
        
        .PARAMETER SplunkIndex
            Splunk Inputs.conf Index. Default: "i_splunk_appl_mssql"

        .PARAMETER SplunkSourcetype
            Splunk Inputs.conf Sourcetype. Default: "mssql_error"

        .NOTES
            Name: Splunk_MSSQL
            Author: Lars Richter
            DateCreated: 13 APRIL 2015

    #>
Param (
    [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
    [Alias('__Server','DNSHostName','IPAddress')]
    [string]$DirApp="C:\Program Files\SplunkUniversalForwarder\etc\apps\mssql",
    [string]$SplunkIndex="i_splunk_appl_mssql",
    [string]$SplunkSourcetype="mssql_error"
    ) 

# Splunk Config File Paths
write-host ("Spluk App Directory: {0} Splunk Index: {1} Splunk Sourcetype: {2}" -f $DirApp, $SplunkIndex, $SplunkSourcetype)
$DirHome= $DirApp + "\default\"
$DirInputs= $DirHome + "inputs.conf"
$DirProps= $DirHome + "props.conf"

Function Get-SQLInstance {  
    <#
        .SYNOPSIS
            Retrieves SQL server information from a local or remote servers.

        .DESCRIPTION
            Retrieves SQL server information from a local or remote servers. Pulls all 
            instances from a SQL server and detects if in a cluster or not.

        .PARAMETER Computername
            Local or remote systems to query for SQL information.

        .NOTES
            Name: Get-SQLInstance
            Author: Boe Prox
            DateCreated: 07 SEPT 2013

        .EXAMPLE
            Get-SQLInstance -Computername DC1

            SQLInstance   : MSSQLSERVER
            Version       : 10.0.1600.22
            isCluster     : False
            Computername  : DC1
            FullName      : DC1
            isClusterNode : False
            Edition       : Enterprise Edition
            ClusterName   : 
            ClusterNodes  : {}
            Caption       : SQL Server 2008

            SQLInstance   : MINASTIRITH
            Version       : 10.0.1600.22
            isCluster     : False
            Computername  : DC1
            FullName      : DC1\MINASTIRITH
            isClusterNode : False
            Edition       : Enterprise Edition
            ClusterName   : 
            ClusterNodes  : {}
            Caption       : SQL Server 2008

            Description
            -----------
            Retrieves the SQL information from DC1
    #>
    [cmdletbinding()] 
    Param (
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('__Server','DNSHostName','IPAddress')]
        [string[]]$ComputerName = $env:COMPUTERNAME
    ) 
    Process {
        Write-Verbose "Get-SQLInstance().."
        ForEach ($Computer in $Computername) {
            $Computer = $computer -replace '(.*?)\..+','$1'
            Write-Verbose ("Checking computer: {0}" -f $Computer)
            Try { 
                $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer) 
                $baseKeys = "SOFTWARE\\Microsoft\\Microsoft SQL Server",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server"
                If ($reg.OpenSubKey($basekeys[0])) {
                    $regPath = $basekeys[0]
                } ElseIf ($reg.OpenSubKey($basekeys[1])) {
                    $regPath = $basekeys[1]
                } Else {
                    Continue
                }
                $regKey= $reg.OpenSubKey("$regPath")
                If ($regKey.GetSubKeyNames() -contains "Instance Names") {
                    $regKey= $reg.OpenSubKey("$regpath\\Instance Names\\SQL" ) 
                    $instances = @($regkey.GetValueNames())
                } ElseIf ($regKey.GetValueNames() -contains 'InstalledInstances') {
                    $isCluster = $False
                    $instances = $regKey.GetValue('InstalledInstances')
                } Else {
                    Continue
                }
                Write-Verbose ("Instances: {0}" -f $instances)
                If ($instances.count -gt 0) { 
                    ForEach ($instance in $instances) {
                        $nodes = New-Object System.Collections.Arraylist
                        $clusterName = $Null
                        $isCluster = $False
                        $instanceValue = $regKey.GetValue($instance)
                        $instanceReg = $reg.OpenSubKey("$regpath\\$instanceValue")
                        If ($instanceReg.GetSubKeyNames() -contains "Cluster") {
                            $isCluster = $True
                            $instanceRegCluster = $instanceReg.OpenSubKey('Cluster')
                            $clusterName = $instanceRegCluster.GetValue('ClusterName')
                            $clusterReg = $reg.OpenSubKey("Cluster\\Nodes")                            
                            $clusterReg.GetSubKeyNames() | ForEach {
                                $null = $nodes.Add($clusterReg.OpenSubKey($_).GetValue('NodeName'))
                            }
                        }
                        $instanceRegSetup = $instanceReg.OpenSubKey("Setup")
                        Try {
                            $edition = $instanceRegSetup.GetValue('Edition')
                        } Catch {
                            $edition = $Null
                        }
                        Write-Verbose ("Instance: {0} instanceValue: {1} isCluster: {2} clusterName: {3}" -f $instance, $instanceValue, $isCluster, $clusterName)
                        Try {
                            $ErrorActionPreference = 'Stop'
                            #Get from filename to determine version
                            $servicesReg = $reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Services")
                            $serviceKey = $servicesReg.GetSubKeyNames() | Where {
                                $_ -match "$instance"
                            } | Select -First 1
                            $service = $servicesReg.OpenSubKey($serviceKey).GetValue('ImagePath')
                            $file = $service -replace '^.*(\w:\\.*\\sqlservr.exe).*','$1'
                            $version = (Get-Item ("\\$Computer\$($file -replace ":","$")")).VersionInfo.ProductVersion
                        } Catch {
                            #Use potentially less accurate version from registry
                            $Version = $instanceRegSetup.GetValue('Version')
                        } Finally {
                            $ErrorActionPreference = 'Continue'
                        }
                        New-Object PSObject -Property @{
                            Computername = $Computer
                            SQLInstance = $instance
                            Edition = $edition
                            Version = $version
                            Caption = {Switch -Regex ($version) {
                                "^14" {'SQL Server 2014';Break}
                                "^11" {'SQL Server 2012';Break}
                                "^10\.5" {'SQL Server 2008 R2';Break}
                                "^10" {'SQL Server 2008';Break}
                                "^9"  {'SQL Server 2005';Break}
                                "^8"  {'SQL Server 2000';Break}
                                Default {'Unknown'}
                            }}.InvokeReturnAsIs()
                            isCluster = $isCluster
                            isClusterNode = ($nodes -contains $Computer)
                            ClusterName = $clusterName
                            ClusterNodes = ($nodes -ne $Computer)
                            FullName = {
                                If ($Instance -eq 'MSSQLSERVER') {
                                    $Computer
                                } Else {
                                    "$($Computer)\$($instance)"
                                }
                            }.InvokeReturnAsIs()
                        }
                    }
                }
            } Catch { 
                Write-Warning ("{0}: {1}" -f $Computer,$_.Exception.Message)
            }  
        }   
    }
} #end function Get-SQLInstance


Function Get-ErrorLogPath
{
    Param([PSObject]$SQLServer = "(local)")
    $instance = ""
    if ($SQLServer.isCluster) { 
        write-verbose ("Use Cluster Name")
        $instance = $SQLServer.ClusterName + "\" + $SQLServer.SQLInstance # Use Clustername + SQLInstance

        $srv = new-object ("Microsoft.SqlServer.Management.Smo.Server") $instance 
        
        # try clustername without instance if logpath is null
        if (-not($srv.Information.ErrorLogPath)) {
            write-verbose ("Get-ErrorLogPath: SQLServer: >{0}< failed. Trying without instancename.." -f $instance)       
            $instance = $SQLServer.ClusterName
            $srv = new-object ("Microsoft.SqlServer.Management.Smo.Server") $instance 
        }

    } else {
        write-verbose ("Use Computer Name")
        if($SQLServer.SQLInstance -eq "MSSQLSERVER") { # default instance
            $instance = "."
        } else {
            $instance = ".\"+$SQLServer.SQLInstance
        }
        $srv = new-object ("Microsoft.SqlServer.Management.Smo.Server") $instance 
    }


    write-host "Info: Get-ErrorLogPath: SQLServer: >" $instance  "< LogPath: >" $srv.Information.ErrorLogPath "<"       
    if (-not($srv.Information.ErrorLogPath)) {
        write-error "Can't get ErrorLogPath! Propably because of missing Assembly or inactive Clusternode!" -RecommendedAction "Install Microsoft.SqlServer.Smo Assembly"
        Exit 1
    }
    $srv.Information.ErrorLogPath
} #end function Get-ErrorLogPath

# Include SMO Assembly
try {
    add-type -AssemblyName "Microsoft.SqlServer.Smo, 
    Version=10.0.0.0,
    Culture=neutral, 
    PublicKeyToken=89845dcd8080cc91" -EA Stop }
catch {
    try {
        add-type -AssemblyName "Microsoft.SqlServer.Smo, 
        Version=11.0.0.0,
        Culture=neutral, 
        PublicKeyToken=89845dcd8080cc91" -EA Stop }
        catch {
            add-type -AssemblyName "Microsoft.SqlServer.Smo"
        }
} #end Include SMO Assembly

$instances = New-Object System.Collections.ArrayList

$logPaths = @()
# Get Default Instance Log Path
# $logPaths+=Get-ErrorLogPath -SQLSERVER "." #TODO: Check if Default Instance is in list when outcommented
$instancesToDelete = New-Object System.Collections.ArrayList    

# Get all instances
#$instances = Get-SQLInstance -ComputerName $Computernames -Verbose
$instances = Get-SQLInstance -Verbose

# iterate through instances
ForEach ($instance in $instances) {
    Write-Verbose ("ForEach instance: {0}" -f $instance)
    $instLogPath=Get-ErrorLogPath -SQLSERVER $instance 
    #Check if Logpath is allready in list
    Write-Verbose ("logPaths: {0}" -f $instLogPath)
    $pathInList = 0
    ForEach ($curLogPath in $logPaths) {
        if ($curLogPath -eq $instLogPath) {
            $pathInList = 1
            $instancesToDelete.add($instance)
        }
    }     
    #Add Path to List
    if (-not $pathInList) {     
        Write-Verbose "Add Path To List: " 
        $logPaths+=$instLogPath
    }
}

# Remove unwanted instances because same ErrorPath as MainInstance
ForEach($instance in $instancesToDelete) {
    $instances.Remove($instance)
    Write-Verbose ("Info: Ignore Instance: {0} because duplicated ErrorPath" -f $instance)
}

# Remove Splunk App Config
Write-Verbose "Remove Old Config Files.."
try{ 
    if(Test-Path $DirHome){ # remove path if exists
        Remove-Item  -Force -confirm:$false -Recurse $DirApp
        write-verbose ("Removed Dir: {0}" -f $DirApp)
    }
}
catch{
    write-error "Cant delete directory: " $DirApp -RecommendedAction "Run script as admin"
}

# Create Folder and Config Files if directy is deleted
Write-Verbose "Write Config Files.." 
if(-NOT (Test-Path $DirHome)){
    mkdir $DirHome
    New-Item -Path $DirHome -Name "inputs.conf" -ItemType File
    New-Item -Path $DirHome -Name "props.conf" -ItemType File
    # Add props.conf Content
    Write-Verbose ("Write Props.conf")
    Add-content $DirProps ("[" + $SplunkSourcetype + "]")
    Add-content $DirProps "SHOULD_LINEMERGE = true "
    Add-content $DirProps "BREAK_ONLY_BEFORE_DATE = true "
    Add-content $DirProps "MAX_TIMESTAMP_LOOKAHEAD = 22 "
    Add-content $DirProps "CHARSET = UTF-16LE"
    Add-content $DirProps "NO_BINARY_CHECK = true"
 
    # Add inputs.conf Content
    Write-Verbose ("Write Inputs.conf")
    $logPathNum = 0
    ForEach ($instance in $instances) {
        Write-Verbose ("Write Instance: {0}" -f $instance)
        # Add Monitor Path
        $monitorPath= "[monitor://"+ $logPaths[$logPathNum]+"\ERRORLOG]"    
        Add-content $DirInputs $monitorPath
        Add-content $DirInputs "disabled=false"          
        Add-content $DirInputs ("index=" + $SplunkIndex)
        Add-content $DirInputs ("sourcetype=" + $SplunkSourcetype)
        Add-content $DirInputs ("source=" + $instance.SQLInstance)
        Add-content $DirInputs "followTail=false"
        #Add host Parameter for Clusters
        if ($instance.isCluster) {
            write-verbose ("isCluster: {0} ClusterName: {1}" -f $instance.isCluster, $instance.ClusterName) 
            Add-content $DirInputs ("host="+$instance.ClusterName)
        }

        #Add Agent
        $monitorPath= "[monitor://"+ $logPaths[$logPathNum]+"\SQLAGENT.OUT]"
        Add-content $DirInputs $monitorPath
        Add-content $DirInputs "disabled=false"          
        Add-content $DirInputs ("index=" + $SplunkIndex)
        Add-content $DirInputs ("sourcetype=" + $SplunkSourcetype)
        Add-content $DirInputs ("source=" + $instance.SQLInstance + "(Agent)")
        Add-content $DirInputs "followTail=false"
        #Add host Parameter for Clusters
        if ($instance.isCluster) {
            Add-content $DirInputs ("host="+$instance.ClusterName)
        }       

        $logPathNum++
    }
}
