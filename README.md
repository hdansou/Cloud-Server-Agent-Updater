# Cloud Server Agent Updater

CloudServerAgentUpdater is a Powershell module that upgrades the Nova Agent to the latest version on Windows Cloud Servers.


## Installation
First, check your Powershell version 
```
$PSVersionTable.PSVersion.Major
```

Next, choose your preferred installation method.

### Any Powershell Version
Choose the module path you want to use to store your module. The following command will list all your module directories.
```
 $env:PSModulePath -split ';'
```
Download and extract the Module in your choosen module directory.
Last, import the CloudServerAgentUpdater module by running:
```
Import-Module  CloudServerAgentUpdater
```


### Powershell version 5
```
Install-Module CloudServerAgentUpdater -Scope CurrentUser
```


## Update
To update the Module to the latest version, you may run the following command if you using Powershell version 5:
```
Update-Module CloudServerAgentUpdater -Force
```

For any powershell version, you may replace the module directory with the latest one. The process is similar to the above installation instuction.
