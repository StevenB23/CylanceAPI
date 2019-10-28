# CylanceAPI
* Cylance API scripts used for automation

Resources:

https://github.com/Maliek/Cylance-API-PS

https://github.com/jan-tee/cycli-examples

## INSTALL POWERHSHELLGET AND CYCLI MODULES 
* First Trust the Powershell Gallery(MS Owned)
* PS Commands to run:
> Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
> install-module -name PowershellGET -Force 
* restart powershell 
> Install-module -name CyCli


## AddDeviceToZone
* Ingests an excel file with a "Machines Name" column to add to a specified zone. The script was modified from the original to connect to the API and perform action in one go. Modify the API TenantID, Auth url and APIID, variables and supply the API secret within the Command line

> .\AddDevicesToZone.ps1 xxxxx-xxxxxx-xxxx MyZone TheMachines.xlsx

