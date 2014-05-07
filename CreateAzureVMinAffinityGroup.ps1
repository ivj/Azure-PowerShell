param (
    # Cloud service name to deploy the VMs to
    [Parameter(Mandatory = $true)]
    [String]$cloudserviceName,
    
    # Name of the Virtual Machine to create
    [Parameter(Mandatory = $true)]
    [String]$VMName,

    # Name of the Instance
    [Parameter(Mandatory = $true)]
    [String]$instancesize,

    # Name of the Instance
    [Parameter(Mandatory = $true)]
    [String]$vmloginname,

    # Name of the Instance
    [Parameter(Mandatory = $true)]
    [String]$vmLoginpwd,

    # Number of disks
    [Parameter(Mandatory = $true)]
    [String]$NoOfDisks,

    # Name of the Disk
    [Parameter(Mandatory = $true)]
    [String]$diskName,

    # Size of disk
    [Parameter(Mandatory = $true)]
    [String]$diskSize
        
    )

$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath

# Add Azure account
#Add-AzureAccount

$subscr = Get-AzureSubscription # Retrieve with Get-AzureSubscription 
$subscriptionName = $subscr[0].SubscriptionName
$subscriptionName

if($subscr.Count -gt 1){
    Write-Host("We have found multiple Azure Subscription for your account, Please select one to proceed.")
    $subsName = Read-Host 'Enter Azure Subscription name.'
    $subscriptionName = $subsName
}

$storage = Get-AzureStorageAccount # Retreive with Get-AzureStorageAccount  
$storageAccountName = $storage.StorageAccountName
$storageAccountName

if($storage.Count -gt 1){
    Write-Host("We have found multiple storage account in your subscrition, Please select one storage account to proceed.")
    $storageName = Read-Host 'Enter Storage account name.'
    #$name
    Set-AzureStorageAccount -StorageAccountName $storageName
    $storageAccountName = $storageName
}

# Specify the storage account location to store the newly created VHDs  
Set-AzureSubscription -SubscriptionName $subscriptionName #-CurrentStorageAccount $storageAccountName  
  
# Select the correct subscription (allows multiple subscription support)  
Select-AzureSubscription -SubscriptionName $subscriptionName  

$affinityGroup = Get-AzureAffinityGroup
$affinityGroup = $affinityGroup.Name
$affinityGroup

if($affinityGroup.Count -ne 0){

    if($affinityGroup.Count -gt 1)
    {
        Write-Host("We have found multiple affinity groups in your subscrition, Please select one to proceed.")
        $agName = Read-Host 'Enter affinity group name.'
        $affinityGroup = $agName
    }
    else{
        $affinityGroup = $affinityGroup.Name
    }
}
else{
    Write-Host("We have not found any affinity groups in your subscrition, Please create one to proceed.")
    $newAGName = Read-Host 'Enter affinity group name.'
    New-AzureAffinityGroup -Name $newAGName -Description "New Affinity Group"
    $affinityGroup = $newAGName
}
 
# Retreive with Get-AzureLocation  
$location = Get-AzureLocation
$location[1].Name

# Retrieve Server 2012 image name with Get-AzureVMImage 

$imageName = Get-AzureVMImage
$imageName = $imageName | 
             Where-Object {( 
                              $_.ImageFamily -ilike "Windows Server 2012*"
                          )}
$imageName = $imageName[0].ImageName
$imageName

$getVNetCommand = $dir + "\GetVNetName.ps1"
invoke-expression -Command $getVNetCommand 

$fileName = "MyAzVNet.xml"
$filePath = $dir + "\" + $fileName

[xml]$vNetName = Get-Content $filePath
$virtualNetworkName = $vNetName.NetworkConfiguration.VirtualNetworkConfiguration.VirtualNetworkSites.VirtualNetworkSite.name
$virtualNetworkName

# ExtraSmall, Small, Medium, Large, ExtraLarge 
$instanceSize = $instancesize  
$instancesize

# function to create new CloudService
# Has to be a unique name. Verify with Test-AzureService 
function CreateCloudService($ServiceName, $AffinityGroup){
    New-AzureService -ServiceName $cloudserviceName -AffinityGroup $AffinityGroup -ErrorAction SilentlyContinue
}

CreateCloudService -ServiceName $cloudserviceName -AffinityGroup $affinityGroup

function AddDataDiskToVM($noOfDisk, $diskSize, $diskName){
    for($i = 0; $i -le $noOfDisk; $i++){
        Add-AzureDataDisk -CreateNew -DiskSizeInGB $diskSize -DiskLabel $diskName -LUN 0
      }
}
 
# Server Name 
$adminname = $vmloginname
$admpwd = $vmLoginpwd

Write-host("Azure VM Creation Started.")

$vm1 = New-AzureVMConfig -Name $VMName -InstanceSize $instanceSize -Image $imageName |
       Add-AzureProvisioningConfig -Windows -AdminUserName $adminname -Password $admpwd |
       Add-AzureDataDisk -CreateNew -DiskSizeInGB $diskSize -DiskLabel $diskName -LUN 0 -HostCaching None |
       New-AzureVM -ServiceName $cloudserviceName -AffinityGroup $affinityGroup `
       -VNetName $virtualNetworkName -ErrorAction SilentlyContinue

Write-host("Azure VM Creation Completed.")

Write-host("Installing Remote Management Certificate from the Virtual Machine started")
# Install a remote management certificate from the Virtual Machine.
$installWinRMCert = $dir + "\InstallWinRMCertificate.ps1 -ServiceName $cloudserviceName -VMName $VMName"
invoke-expression -Command $installWinRMCert
#Install-WinRmCertificate -ServiceName $cloudserviceName -VMName $VMName

Write-host("Installing Remote Management Certificate from the Virtual Machine Completed")

$winRmUri = Get-AzureWinRMUri -ServiceName $cloudserviceName -Name $VMName

Write-host("Configuring Virtual Machine Started")

$acl1 = New-AzureAclConfig
Set-AzureAclConfig –AddRule –ACL $acl1 –Order 100 –Action permit –RemoteSubnet “192.168.11.0/24” –Description “Remote Desktop ACL config” 
 
Get-AzureVM –ServiceName $cloudserviceName –Name $VMName | Set-AzureEndpoint –Name “RDP” –Protocol tcp –Localport 3389 -PublicPort 3389 –ACL $acl1 | Update-AzureVM 

Write-host("Configuring Virtual Machine Completed")

Remove-Item $filePath
