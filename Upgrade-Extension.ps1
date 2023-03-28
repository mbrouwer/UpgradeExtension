[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)][switch]$AllowCustomImage,    
    [Parameter(Mandatory = $false)][string]$VMFilter,
    [Parameter(Mandatory = $false)][string]$SubscriptionFilter,
    [Parameter(Mandatory = $false)][switch]$InstallExtension,
    [Parameter(Mandatory = $false)][switch]$NoOutput,
    [Parameter(Mandatory = $false)][string]$extensionPublisherName = "Microsoft.GuestConfiguration",
    [Parameter(Mandatory = $false)][string]$extensionTypeName = "ConfigurationforWindows",
    [Parameter(Mandatory = $false)][string]$extensionName = "AzurePolicyforWindows",
    [Parameter(Mandatory = $false)][string]$extensionTypeVersion = "1.1",
    [Parameter(Mandatory = $false)][switch]$enableAutomaticUpgrade = $false,
    [Parameter(Mandatory = $false)][switch]$autoUpgradeMinorVersion = $true
    
)

Function Search-Azure {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$Query
    )
    $url = "https://management.azure.com/providers/Microsoft.ResourceGraph/resources?api-version=2021-03-01"
    $body = [PSCustomObject]@{
        query = $query
    } | ConvertTo-Json

    $return = @()

    $graphResult = (Invoke-AzRestMethod -Uri $url -Method POST -Payload $body).Content | ConvertFrom-Json -Depth 99
    $return += $graphResult.data
    if ($graphResult.'$skipToken') {
        do {
            $body = [PSCustomObject]@{
                query   = $query
                options = [PSCustomObject]@{
                    '$skipToken' = $graphResult.'$skipToken'
                }
            } | ConvertTo-Json
    
            $graphResult = (Invoke-AzRestMethod -Uri $url -Method POST -Payload $body).Content | ConvertFrom-Json -Depth 99
            $return += $graphResult.data
        } while (
            $graphResult.'$skipToken'
        )
    }
    return $return
}
function Install-VMExtension {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$virtualMachineId
    )
    $url = "https://management.azure.com$($virtualMachineId)/extensions/$extensionName?api-version=2022-03-01"

    $body = [PSCustomObject]@{
        properties = [PSCustomObject]@{
            publisher               = $extensionPublisherName
            type                    = $extensionTypeName
            typeHandlerVersion      = $extensionTypeVersion
            autoUpgradeMinorVersion = $autoUpgradeMinorVersion
            enableAutomaticUpgrade  = $enableAutomaticUpgrade
        }
        location   = "westeurope"
    } | ConvertTo-Json -Depth 10
        
    $return = Invoke-AzRestMethod -Uri $url -Payload $body -Method PUT
    return $return
}

$startTime = Get-Date
$InformationPreference = "Continue"

$return = @()
$virtualMachines = @()


Write-Information "Getting all VMs"
$virtualMachines = Search-Azure -Query "resources | where type == 'microsoft.compute/virtualmachines' | mv-expand customImage = properties.storageProfile.imageReference.id"
$extensions = Search-Azure -Query "resources | where type == 'microsoft.compute/virtualmachines/extensions' | where name == '$($extensionName)'"

Write-Information "Getting latest extension version"
$url = "https://management.azure.com/subscriptions/$((Get-AzContext).Subscription.Id)/providers/Microsoft.Compute/locations/westeurope/publishers/$($extensionPublisherName)/artifacttypes/vmextension/types/$($extensionTypeName)/versions?api-version=2022-03-01"
$latestVersion = ((Invoke-AzRestMethod -Uri $url).Content | ConvertFrom-Json | Select-Object @{Name = 'Version'; Expression = { [system.version]$_.name } } | Sort-Object Version -Descending)[0].Version.ToString()
Write-Information "Latest version is $($latestVersion)"

if (-not [string]::IsNullOrEmpty($vmFilter)) {
    $virtualMachines = ($virtualMachines | Where-Object { $_.Name -eq "$($vmFilter)" })    
}

if (-not [string]::IsNullOrEmpty($SubscriptionFilter)) {
    $virtualMachines = $virtualMachines | Where-Object { $_.SubscriptionId -eq $SubscriptionFilter }
}

if (-not $AllowCustomImage) {
    $virtualMachines = $virtualMachines | Where-Object { -not $_.customImage }
}

# $ignoredOS = @("Canonical", "redhat", "ubuntu", "centos", "zscaler", "center-for-internet-security-inc")
# $virtualMachines = $virtualMachines | Where-Object { $_.Properties.storageProfile.imageReference.publisher -notin $ignoredOS -and $_.properties.extended.instanceView.osName -notin $ignoredOS}
$virtualMachines = $virtualMachines | Where-Object { $_.properties.storageProfile.osDisk.osType -ne "Linux" }

$return = @()

Write-Information "Parsing extensions for $($virtualMachines.count) VM(s)"
$vmCount = 0
foreach ($virtualMachine in $virtualMachines) {
    $vmCount++
    # Write-Progress -Activity "Parsing VM '$($virtualMachine.Name)'" -PercentComplete (($vmCount / $virtualMachines.count)*100)

    if (@($extensions | Where-Object { $_.id -like "$($virtualMachine.id)*" }).Count -gt 0) {
        $extensionUrl = "https://management.azure.com$($virtualMachine.Id)/extensions/$($extensionName)?`$expand=instanceView&api-version=2021-11-01"
        $extensionObject = ((Invoke-AzRestMethod -Uri $extensionUrl).Content | ConvertFrom-Json -Depth 99 | Where-Object { $null -ne $_.id })
    }
    else {
        $extensionObject = $null
    }

    $powerState = ($virtualMachine.properties.extended.instanceView.powerState | Where-Object { $_.code -like "Powerstate/*" }).code.split("/")[-1]

    $returnObject = [PSCustomObject]@{
        VMName             = $virtualMachine.name
        SubscriptionId     = $virtualMachine.subscriptionId
        OSName             = $virtualMachine.properties.extended.instanceView.osName
        OSVersion          = $virtualMachine.properties.extended.instanceView.osVersion
        CustomImage        = $virtualMachine.properties.storageProfile.imageReference.id ? $true : $false
        ImageReference     = $virtualMachine.properties.storageProfile.imageReference
        PoweredOn          = $powerState -eq "running" ? $true : $false
        PowerState         = ($virtualMachine.properties.extended.instanceView.powerState | Where-Object { $_.code -like "Powerstate/*" }).code.split("/")[ - 1]
        ExtensionInstalled = ($extensionObject.count -gt 0)
        ExtensionVersion   = $extensionObject.count -gt 0 ? $extensionObject.properties.instanceView.typeHandlerVersion : $null
        AutomaticUpgrade   = $extensionObject.count -gt 0 ? $extensionObject.properties.enableAutomaticUpgrade : $null
        Identity           = $virtualMachine.identity
        UpgradeExtension   = $false
        UpgradeResult      = $null
        VMObject           = $virtualMachine
        # ExtensionObject    = $extensionObject

    }

    if ($returnObject.ExtensionInstalled) {
        if ($returnObject.PoweredOn) {
            if (-not ([System.Version]$returnObject.ExtensionVersion -ge [System.Version]$latestVersion)) {
                $returnObject.UpgradeExtension = $true
            }

            if (-not $returnObject.AutomaticUpgrade) {
                $returnObject.UpgradeExtension = $true
            }
        }
    }
    else {
        if ($returnObject.PoweredOn) {
            $returnObject.UpgradeExtension = $true
        }
    }

    if ($returnObject.UpgradeExtension -and $installExtension) {
        Write-Information "Installing extension on VM '$($returnObject.VMName)'"
        $returnObject.UpgradeResult = (Install-VMExtension -virtualMachineId $virtualMachine.Id)
    }

    $return += $returnObject
}

$stopTime = Get-Date
Write-Information "Total Time : $((New-TimeSpan -Start $startTime -End $stopTime).TotalSeconds) seconds"

if (-not $NoOutput) {
    return $return
}
