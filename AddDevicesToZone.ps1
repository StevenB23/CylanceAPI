<#
.DESCRIPTION
    This tool will take an Excel file, look at its "Machine Name" column, and add all of the hosts in it that already exist
    in the console to a specific zone.

    This is useful in situations where you need to add arbitrary groups not based on some criteria that can be expressed via zone rules.

    The Excel file needs to have a column "Machine Name" on the active worksheet, and this has to contain the device names to add to the zone.
    
    Make sure to specify the entire path to your excel file to avoid errors

#>


[CmdletBinding()]
Param (
    [Parameter(Mandatory=$true)]
    [String]$Secret,
    [Parameter(Mandatory=$true)]
    [String]$ZoneName,
    [Parameter(Mandatory=$true)]
    [String]$ExcelFile
    )

Import-Module CyCLI
Import-Module ImportExcel

# Making the API connection
$APIId = "3dx3d3f5-dd21-x3x3dx-x3x3x-33d3dc4d4d43d" # enter the API ID
#$APIsecret = ConvertTo-SecureString -String "14741a0b-c1a7-43ef-899b-b444c99003ef" -AsPlainText -Force
$APIsecret = ConvertTo-SecureString -String $Secret -AsPlainText -Force
$TenantId = "33243dxfd43d-xxxxxxx-f343ggg4411" # enter your tenant ID
$auth = "https://protectapi<-region>.cylance.com/auth/v2/token" #Enter API authentication URL based on your region
Get-CyAPI -APIId $APIId -APISecret $APIsecret -APITenantId $TenantId -APIAuthUrl $auth

# Creates zone if it does not exist
$Zone = Get-CyZone -Name $ZoneName
if ($Zone -eq $null) {
    $Zone = New-CyZone -Name $ZoneName -Criticality Normal
}

# Get list of devices to add to zone
$DevicesToAdd = @( Import-Excel -Path $ExcelFile | Select-Object "Machine Name")

# Identify devices that already exist in tenant 
Write-Host -NoNewline "There were $($DevicesToAdd.Count) devices in the Excel file, of which "
$ExistingDevices = @( Get-CyDeviceList | Where-Object { $DevicesToAdd."Machine Name" -Contains $_.name } )
Write-Host "$($ExistingDevices.Count) devices exist in the tenant."

# Add those devices to zone
Write-Host -NoNewline "Adding devices to the zone $($Zone.name)..."
$ExistingDevices | Add-CyDeviceToZone -Zone $Zone
Write-Host "done."
