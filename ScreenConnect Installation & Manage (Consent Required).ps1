# This is an amalgamation of Steven Grabowski's & Sam James' scripts. Thanks Steven & Sam!

# This will deploy Screenconnect to your Syncro asset also enabling syncro integration.
#   It will set the asset friendly name and company name from syncro in your screenconnect portal. (currently, the official syncro integration dosn't do this)
#   It will integrate screenconnect with syncro to allow starting a remote session from the Syncro asset page. (Just like the official integration)

# When creating this in syncro, create two Syncro platform variables...
#   Variable Name - $FriendlyName / Variable Type - platform / Value - asset_custom_field_device_name
#   Variable Name - $CompanyName / Variable Type - platform / Value - customer_business_name_or_customer_full_name


Import-Module $env:SyncroModule

# Properly encode URL parameters
$EncodedCompanyName = [uri]::EscapeDataString($CompanyName)
$EncodedFriendlyName = [uri]::EscapeDataString($FriendlyName)

# URL for your screenconnect instance, you can customize the port or leave it off
$scdomain = "bullertech.screenconnect.com"

# URL for ScreenConnect msi download
# This is the URL for the ScreenConnect msi download. It includes the encoded company name and friendly name.
# The Fifth parameter is True, which means the user will be prompted for consent.
$url = "https://$scdomain/Bin/ScreenConnect.ClientSetup.msi?e=Access&y=Guest&t=$EncodedFriendlyName&c=$EncodedCompanyName&c=&c=&c=True&c=&c=&c=&c="

# put your instance string here find in add/remove programs
$scinstance = "c4d53e2bd6ff64ec"

# put your instance string here find in add/remove programs
$serviceName = "ScreenConnect Client ($scinstance)"

# Your syncro subdomain
$subdomain = "bullertech"

If (Get-Service $serviceName -ErrorAction SilentlyContinue) {
   If ((Get-Service $serviceName).Status -eq 'Running') {
       Write-Host "$serviceName is running, skipping service start or install, and performing syncro GUID match."
   } Else {
       #$ticket = Create-Syncro-Ticket -Subdomain $subdomain -Subject "ScreenConnect needs started on $env:computername" -IssueType "Automated" -Status "New"
       Write-Host "$serviceName found, but it is not running for some reason."
       Write-Host "starting $servicename"
       start-service $serviceName
       #$startAt = (Get-Date).AddMinutes(-30).toString("o")
       #Create-Syncro-Ticket-TimerEntry -Subdomain $subdomain -TicketIdOrNumber $ticket.ticket.id -StartTime $startAt -DurationMinutes 5 -Notes "Automated system cleaned up the disk space." -UserIdOrEmail $yourEmail
        
        If ((Get-Service $serviceName).Status -eq 'Running') {
           Write-Host "$serviceName started after manual start commant, now performing syncro GUID match."
        } Else {
            Write-Host "$serviceName not responding to start commant, performing uninstall and fresh install, then performing syncro GUID match."
            # Use modern PowerShell method to uninstall ScreenConnect
            try {
                $product = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*ScreenConnect Client*" -and $_.Name -like "*$scinstance*" }
                if ($product) {
                    Write-Host "Uninstalling existing ScreenConnect installation..."
                    $result = $product.Uninstall()
                    if ($result.ReturnValue -eq 0) {
                        Write-Host "Successfully uninstalled existing ScreenConnect installation"
                    } else {
                        Write-Host "Warning: Uninstall returned error code: $($result.ReturnValue)"
                    }
                }
            } catch {
                Write-Host "Warning: Failed to uninstall existing installation: $($_.Exception.Message)"
            }
            
            Write-Host "$serviceName not found - need to install"
            (new-object System.Net.WebClient).DownloadFile($url,'C:\windows\temp\sc.msi')
            msiexec.exe /i c:\windows\temp\sc.msi /quiet
            Write-Host "$serviceName install command sent."
            
        }
   }
} Else {
   Write-Host "$serviceName not found - need to install"
   (new-object System.Net.WebClient).DownloadFile($url,'C:\windows\temp\sc.msi')
   msiexec.exe /i c:\windows\temp\sc.msi /quiet
}

Start-Sleep -Seconds 10

#Get the Screenconnect GUID and write to Syncro asset.
$val = (Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Services\$serviceName").ImagePath
$Regex = [Regex]::new("(?<=s=)(.*?)(?=&)")
$Match = $Regex.Match($val)
if($Match.Success)
{
    Set-Asset-Field -Subdomain $subdomain -Name "ScreenConnect GUID" -Value $Match
    $ScreenConnectUrl = "https://$scdomain/Host#Access/All%20Machines//$Match/Join"
    Write-Host ScreenConnect URL Is: $ScreenConnectUrl
}