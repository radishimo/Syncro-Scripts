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
            # Comprehensive ScreenConnect uninstall using multiple methods
            Write-Host "Performing comprehensive ScreenConnect uninstall..."
            
            # Method 1: Try using Get-Package and Uninstall-Package (PowerShell 5.1+)
            try {
                Write-Host "Attempting uninstall using PowerShell Package Management..."
                $package = Get-Package -Name "*ScreenConnect*" -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*$scinstance*" }
                if ($package) {
                    Uninstall-Package -Name $package.Name -Force -ErrorAction Stop
                    Write-Host "Successfully uninstalled using PowerShell Package Management"
                    Write-Host "Uninstall completed, proceeding with fresh install..."
                    # Skip to fresh install - don't try other methods
                    $uninstallSuccess = $true
                }
            } catch {
                Write-Host "PowerShell Package Management method failed: $($_.Exception.Message)"
            }

            # Method 2: Try using WMI with PowerShell (more reliable than wmic command)
            if (-not $uninstallSuccess) {
                try {
                    Write-Host "Attempting uninstall using WMI via PowerShell..."
                    $product = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*ScreenConnect Client*" -and $_.Name -like "*$scinstance*" }
                    if ($product) {
                        Write-Host "Found ScreenConnect installation, uninstalling..."
                        $result = $product.Uninstall()
                        if ($result.ReturnValue -eq 0) {
                            Write-Host "Successfully uninstalled using WMI"
                            Write-Host "Uninstall completed, proceeding with fresh install..."
                            # Skip to fresh install - don't try other methods
                            $uninstallSuccess = $true
                        } else {
                            Write-Host "WMI uninstall returned error code: $($result.ReturnValue)"
                        }
                    } else {
                        Write-Host "ScreenConnect installation not found via WMI"
                    }
                } catch {
                    Write-Host "WMI method failed: $($_.Exception.Message)"
                }
            }

            # Method 3: Try using MSI directly if we can find the product code
            if (-not $uninstallSuccess) {
                try {
                    Write-Host "Attempting uninstall using MSI product code..."
                    $msiProducts = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*ScreenConnect*" }
                    foreach ($product in $msiProducts) {
                        if ($product.Name -like "*$scinstance*") {
                            Write-Host "Found MSI product: $($product.Name)"
                            $result = $product.Uninstall()
                            if ($result.ReturnValue -eq 0) {
                                Write-Host "Successfully uninstalled using MSI product code"
                                Write-Host "Uninstall completed, proceeding with fresh install..."
                                # Skip to fresh install - don't try other methods
                                $uninstallSuccess = $true
                                break
                            }
                        }
                    }
                } catch {
                    Write-Host "MSI product code method failed: $($_.Exception.Message)"
                }
            }

            # Method 4: Try using Add/Remove Programs registry method
            if (-not $uninstallSuccess) {
                try {
                    Write-Host "Attempting uninstall using registry-based method..."
                    $uninstallKeys = @(
                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
                    )
                    
                    foreach ($keyPath in $uninstallKeys) {
                        $products = Get-ItemProperty $keyPath -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*ScreenConnect*" -and $_.DisplayName -like "*$scinstance*" }
                        foreach ($product in $products) {
                            if ($product.UninstallString) {
                                Write-Host "Found uninstall string: $($product.UninstallString)"
                                $uninstallCmd = $product.UninstallString -replace '"', '' -split ' '
                                $uninstaller = $uninstallCmd[0]
                                $uninstallArgs = ($uninstallCmd[1..($uninstallCmd.Length-1)] -join ' ') + " /quiet /norestart"
                                
                                Write-Host "Running: $uninstaller $uninstallArgs"
                                Start-Process -FilePath $uninstaller -ArgumentList $uninstallArgs -Wait -NoNewWindow
                                Write-Host "Registry-based uninstall completed"
                                Write-Host "Uninstall completed, proceeding with fresh install..."
                                # Skip to fresh install - don't try other methods
                                $uninstallSuccess = $true
                                break
                            }
                        }
                        if ($uninstallSuccess) { break }
                    }
                } catch {
                    Write-Host "Registry-based method failed: $($_.Exception.Message)"
                }
            }
            
            Write-Host "Uninstall process completed, proceeding with fresh install..."
            
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