#requires -Version 1


#region Variables
$BaseUrl = 'https://github.com/rackerlabs/openstack-guest-agents-windows-xenserver/releases/download/'

$NovaAgentZipUrl = $BaseUrl + $NovaAgentVersion.Latest +'/'+ $NovaAgentZip 
$NovaAgentUpdaterZipUrl = $BaseUrl + $NovaAgentVersion.Latest + '/' + $NovaAgentUpdaterZip
$NovaAgentVersion = @{
    'Latest' = '1.3.0.3'
    'Previous' = @{
        '1.3.0.2' = '1.3.0.2'
        '1.3.0.1' = '1.3.0.1'
        '1.2.9.0' = '1.2.9.0'
        '1.2.8.1' = '1.2.8.1'
    }
}


$NovaAgentZip = 'AgentService.zip'
$NovaAgentUpdaterZip = 'UpdateService.zip'
$TempDir = 'C:\Windows\Temp'
$NovaAgentDir = 'C:\Program Files\Rackspace\Cloud Servers\'
$NovaAgentService = @('RackspaceCloudServersAgent', 'RackspaceCloudServersAgentUpdater')


#endregion

#region Functions
function Invoke-Unzip
{
    Param(
        $ZipFile,
        $Destination    
    )
    try
    {
        if(-not (Test-Path -Path $Destination))
        {
            New-Item -ItemType directory -Path $Destination -Force
        } 

        $sh = New-Object -ComObject shell.application
        $sh.namespace($Destination).Copyhere($sh.namespace($ZipFile).items())
    }
    catch [system.Exception] 
    {
        Write-Output -InputObject "[$(Get-Date)] Error  :: Unzipping $ZipFile"
        Write-Output -InputObject "[$(Get-Date)] Details:: $_"
    }
}

Function Invoke-FileDowload 
{
    Param(
        $Url,
        $localpath,
        $Filename
    )
    if(!(Test-Path -Path $localpath)) 
    {
        New-Item -Path $localpath -type directory > $null
    }

    $webclient = New-Object -TypeName System.Net.WebClient

    try 
    {
        $webclient.DownloadFile($Url, $localpath + '\' + $Filename)
        Write-Output -InputObject "[$(Get-Date)] Status  :: Downloaded Successfully $Filename in  $localpath"
    }
    catch [system.Exception] 
    {
        Write-Output -InputObject "[$(Get-Date)] Error  :: Download Failed for $Filename in  $localpath"
        Write-Output -InputObject "[$(Get-Date)] Details:: $_"
    }
}

function Get-ServiceVersion
{
    Param(
        [Array]$Name = 'RackspaceCloudServersAgent'
    )
    try
    {
        $output = New-Object -TypeName System.Collections.Specialized.OrderedDictionary
        foreach ($ServiceName in $Name)
        {
            $ServiceExecutable = (Get-WmiObject -Class win32_Service | Where-Object -FilterScript {
                    $_.Name -contains $ServiceName
                }
            ).PathName 
            $ServiceVersion = (Get-ChildItem -Path $ServiceExecutable).VersionInfo.ProductVersion
            
            $output.Add($ServiceName, $ServiceVersion)     
        }
        $output
    }
    catch
    {
        Write-Output -InputObject "[$(Get-Date)] Error  :: Error when collecting the version $ServiceName"
        Write-Output -InputObject "[$(Get-Date)] Details:: $_"
    }
}

function Test-NovaAgentVersion
{
    Param(
        $VersionLatest
    )
    
    $VersionInstalled = Get-ServiceVersion -Name RackspaceCloudServersAgent
    if ($VersionLatest -gt $VersionInstalled.RackspaceCloudServersAgent)
    {
       [pscustomobject]@{Upgrade = $true}
    }
    else
    {
        [pscustomobject]@{Upgrade =$false}
    }
}

function Update-NovaAgent
{
    Param(
        $LatestNovaAgentVersion = $NovaAgentVersion.latest
    )

    $NovaAgentVersionInstalled = Get-ServiceVersion -Name $NovaAgentService

    if ( -not $(Test-NovaAgentVersion -VersionLatest $NovaAgentVersion.Latest).Upgrade )
    {
        Write-Output -InputObject "[$(Get-Date)] Status  :: Nova Agent Version $($NovaAgentVersionInstalled.RackspaceCloudServersAgent)"
        Write-Output -InputObject "[$(Get-Date)] Status  :: Nova Agent Updater Version $($NovaAgentVersionInstalled.RackspaceCloudServersAgentUpdater)"   
    }
    else
    {
        if((Get-Service -Name $NovaAgentService).Status -ne 'Stopped')
        { 
            Stop-Service -Name $NovaAgentService -Force -ErrorAction SilentlyContinue
            Write-Output -InputObject "[$(Get-Date)] Status  :: $NovaAgentService is stopped"
        }

        Write-Output -InputObject "[$(Get-Date)] Status  :: Downloading the $NovaAgentZip"
        Invoke-FileDowload -Url $NovaAgentZipUrl -localpath $TempDir -Filename $NovaAgentZip

        Write-Output -InputObject "[$(Get-Date)] Status  :: Downloading the $NovaAgentUpdaterZip"
        Invoke-FileDowload -Url $NovaAgentUpdaterZipUrl  -localpath $TempDir -Filename $NovaAgentUpdaterZip

        if (-not(Test-Path -Path (Join-Path -Path $NovaAgentDir -ChildPath $('Agent' + $($NovaAgentVersionInstalled.RackspaceCloudServersAgent)))))
        {
            Write-Output -InputObject "[$(Get-Date)] Status  :: Renaming Agent to $($NovaAgentVersionInstalled.RackspaceCloudServersAgent)"
            Rename-Item -Path (Join-Path -Path $NovaAgentDir -ChildPath 'Agent')  -NewName (
                Join-Path -Path $NovaAgentDir -ChildPath $('Agent' + $($NovaAgentVersionInstalled.RackspaceCloudServersAgent))
            ) -Force

            Write-Output -InputObject "[$(Get-Date)] Status  :: Renaming AgentUpdater to $($NovaAgentVersionInstalled.RackspaceCloudServersAgentUpdater)"
            Rename-Item -Path (Join-Path -Path $NovaAgentDir -ChildPath 'AgentUpdater') -NewName (
                Join-Path -Path $NovaAgentDir -ChildPath $('AgentUpdater' + $($NovaAgentVersionInstalled.RackspaceCloudServersAgentUpdater))
            ) -Force

            Write-Output -InputObject "[$(Get-Date)] Status  :: Unzipping AgentService.zip to Agent)"
            Invoke-Unzip -ZipFile (Join-Path -Path $TempDir -ChildPath 'AgentService.zip') -Destination (
                Join-Path -Path $NovaAgentDir -ChildPath 'Agent'
            )
        
            Write-Output -InputObject "[$(Get-Date)] Status  :: Unzipping UpdateService.zip to AgentUpdater"   
            Invoke-Unzip -ZipFile (Join-Path -Path $TempDir -ChildPath 'UpdateService.zip') -Destination (
                Join-Path -Path $NovaAgentDir -ChildPath 'AgentUpdater'
            )

            Write-Output -InputObject "[$(Get-Date)] Status  :: Cloning the AgentLog from $('Agent' + $($NovaAgentVersionInstalled.RackspaceCloudServersAgent)) to Agent"
            Copy-Item -Path (Join-Path -Path $NovaAgentDir -ChildPath $('Agent' + $($NovaAgentVersionInstalled.RackspaceCloudServersAgent) + "\Agentlog.txt")) -Destination (
                Join-Path -Path $NovaAgentDir -ChildPath $('Agent\Agentlog.txt')) -Force

        }
    
        Write-Output -InputObject "[$(Get-Date)] Status  :: Removing UpdateService.zip to AgentUpdater.zip"
        Remove-Item -Path (Join-Path -Path $TempDir -ChildPath 'AgentService.zip') -Force
        Remove-Item -Path (Join-Path -Path $TempDir -ChildPath 'UpdateService.zip') -Force

        Write-Output -InputObject "[$(Get-Date)] Status  :: Restarting the Agent and AgentUpdater services"
        Start-Service -Name $NovaAgentService -ErrorAction SilentlyContinue -Verbose
    
        if((Get-Service -Name $NovaAgentService).Status -ne 'Running')
        { 
            Restart-Service -Name $NovaAgentService -Force -ErrorAction SilentlyContinue
            Write-Output -InputObject "[$(Get-Date)] Status  :: $NovaAgentService is stopped"
        }
    }
}
#endregion Function



#region MAIN
Update-NovaAgent -LatestNovaAgentVersion $NovaAgentVersion.Latest
#endregion MAIN
