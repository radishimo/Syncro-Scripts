# ScreenConnect Uninstall Script - Windows 11 Compatible
# This script uses modern PowerShell methods to uninstall ScreenConnect

$scinstance = "c4d53e2bd6ff64ec"
$uninstallSuccess = $false

Write-Host "Starting ScreenConnect uninstall process..."

# Method 1: Try using Get-Package and Uninstall-Package (PowerShell 5.1+)
try {
    Write-Host "Attempting uninstall using PowerShell Package Management..."
    $package = Get-Package -Name "*ScreenConnect*" -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*$scinstance*" }
    if ($package) {
        Uninstall-Package -Name $package.Name -Force -ErrorAction Stop
        Write-Host "Successfully uninstalled using PowerShell Package Management"
        $uninstallSuccess = $true
        exit 0
    }
} catch {
    Write-Host "PowerShell Package Management method failed: $($_.Exception.Message)"
}

# Method 2: Try using WMI with PowerShell (more reliable than wmic command)
try {
    Write-Host "Attempting uninstall using WMI via PowerShell..."
    $product = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*ScreenConnect Client*" -and $_.Name -like "*$scinstance*" }
    if ($product) {
        Write-Host "Found ScreenConnect installation, uninstalling..."
        $result = $product.Uninstall()
        if ($result.ReturnValue -eq 0) {
            Write-Host "Successfully uninstalled using WMI"
            $uninstallSuccess = $true
            exit 0
        } else {
            Write-Host "WMI uninstall returned error code: $($result.ReturnValue)"
        }
    } else {
        Write-Host "ScreenConnect installation not found via WMI"
    }
} catch {
    Write-Host "WMI method failed: $($_.Exception.Message)"
}

# Method 3: Try using MSI directly if we can find the product code
try {
    Write-Host "Attempting uninstall using MSI product code..."
    $msiProducts = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*ScreenConnect*" }
    foreach ($product in $msiProducts) {
        if ($product.Name -like "*$scinstance*") {
            Write-Host "Found MSI product: $($product.Name)"
            $result = $product.Uninstall()
            if ($result.ReturnValue -eq 0) {
                Write-Host "Successfully uninstalled using MSI product code"
                $uninstallSuccess = $true
                exit 0
            }
        }
    }
} catch {
    Write-Host "MSI product code method failed: $($_.Exception.Message)"
}

# Method 4: Try using Add/Remove Programs registry method
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
                $uninstallSuccess = $true
                exit 0
            }
        }
    }
} catch {
    Write-Host "Registry-based method failed: $($_.Exception.Message)"
}

# Only show failure messages if no method succeeded
if (-not $uninstallSuccess) {
    Write-Host "ScreenConnect uninstall process completed, but no method succeeded."
    Write-Host "Please verify the installation has been removed. If ScreenConnect is still present, manual removal may be required."
}