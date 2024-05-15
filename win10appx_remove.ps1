function Remove-WinApps {

$appxRaw=(Get-AppxPackage -AllUsers|where {$_.isframework -eq $false}).name
$exceptions=@(
"Microsoft.NET.Native.Framework",
"Microsoft.NET.Native.Runtime",
"Microsoft.AAD.BrokerPlugin",
"Microsoft.Windows.CloudExperienceHost",
"Microsoft.Windows.ShellExperienceHost",
"windows.immersivecontrolpanel",
"Microsoft.BioEnrollment",
"Microsoft.AccountsControl",
"Microsoft.LockApp",
"Microsoft.Windows.AssignedAccessLockApp",
"Microsoft.Windows.ContentDeliveryManager",
"Microsoft.Windows.ParentalControls",
"Microsoft.WindowsFeedback",
"Windows.ContactSupport",
"Windows.PrintDialog",
"Windows.PurchaseDialog",
"windows.devicesflow",
"Microsoft.Windows.SecondaryTileExperience",
"Microsoft.Windows.FeatureOnDemand.InsiderHub",
"Microsoft.XboxGameCallableUI",
"Microsoft.XboxIdentityProvider",
"Windows.MiracastView",
"Microsoft.Advertising.Xaml",
"Microsoft.Services.Store.Engagement",
"Microsoft.VCLibs",
"Microsoft.Windows.Cortana",
"Microsoft.WindowsCalculator",
"Microsoft.Windows.Photos",
"Microsoft.MicrosoftEdge",
"Microsoft.WindowsStore",
"Microsoft.WindowsCamera",
"Microsoft.MSPaint"
)

$appx=@()


foreach ($appRaw in $appxRaw) {
    
    $isException=0

    foreach ($ex in $exceptions){

        if ($appRaw -like "*$ex*") {

        $isException+=1

        }
    }

    if ($isException -eq 0) {

        $appx+=$appRaw
    }
}

foreach ($app in $appx) {
$PackageFullName = (Get-AppxPackage $App -allusers).PackageFullName
 $ProPackageFullName = (Get-AppxProvisionedPackage -online | where {$_.Displayname -eq $App}).PackageName
 write-host $PackageFullName
 Write-Host $ProPackageFullName
 if ($PackageFullName)
 {
 Write-Host "Removing Package: $App"
 remove-AppxPackage -package $PackageFullName
 }
 else
 {
 Write-Host "Unable to find package: $App"
 }
 if ($ProPackageFullName)
 {
 Write-Host "Removing Provisioned Package: $ProPackageFullName"
 Remove-AppxProvisionedPackage -online -packagename $ProPackageFullName
 }
 else
 {
 Write-Host "Unable to find provisioned package: $App"
 }
 }


}

Remove-WinApps
Read-Host "Press ENTER to close..."
