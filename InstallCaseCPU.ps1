# Check if running as administrator
Function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Relaunch script with elevated privileges
If (-Not (Test-Admin)) {
    Write-Host "This script requires administrative privileges. Restarting with elevated permissions..."
    Start-Process powershell "-File $PSCommandPath" -Verb RunAs
    Exit
}

# Extension ID - will be generated when extension is published
# For local testing, use the actual extension ID from chrome://extensions
$global:extensionId = "elhjfifcippogjljafhhginidkbliimn"

# Extension settings JSON - configure update URL based on your hosting
$global:extSettingsValue = @"
{
    "*": {
        "install_sources": [
            "*://raw.githubusercontent.com/*",
            "*://github.com/*"
        ]
    },
    "$global:extensionId": {
        "installation_mode": "force_installed",
        "update_url": "https://raw.githubusercontent.com/Johnny-spec/CaseProEdgeExtension/main/updates.xml",
        "toolbar_state": "force_shown"
    }
}
"@

$global:edgeKey = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
$global:chromeKey = "HKLM:\SOFTWARE\Policies\Google\Chrome"

# Menu function
Function Show-Menu {
    Write-Host "`n=== CaseCPU Extension Manager ===" -ForegroundColor Cyan
    Write-Host "1. Install CaseCPU extension"
    Write-Host "2. Uninstall CaseCPU extension"
    Write-Host "3. Show current installation status"
    Write-Host "4. Exit"
    $choice = Read-Host "Enter your choice"
    Return $choice
}

# Add extension using ExtensionSettings
Function Add-Extension1($baseKey, $browserName) {
    If (-Not (Get-ItemProperty -Path $baseKey -Name "ExtensionSettings" -ErrorAction SilentlyContinue)) {
        New-ItemProperty -Path $baseKey -Name "ExtensionSettings" -Value $global:extSettingsValue -PropertyType String -Force
        Write-Host "Extension installed for $browserName using ExtensionSettings." -ForegroundColor Green
        return $true
    }
    Else {
        # ExtensionSettings exists, process its value
        $currentSetting = Get-ItemProperty -Path $baseKey -Name "ExtensionSettings" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ExtensionSettings
        If ($currentSetting) {
            $extSettingsJson = $currentSetting | ConvertFrom-Json
            If ($extSettingsJson.$global:extensionId -and $extSettingsJson.$global:extensionId.installation_mode -eq "force_installed") {
                Write-Host "Extension is already installed for $browserName using ExtensionSettings." -ForegroundColor Yellow
                return $true
            }
        }
        return $false
    }
}

# Add extension to registry using ExtensionInstallForcelist
Function Add-Extension2($baseKey, $extensionString, $browserName) {
    If (-Not (Test-Path $baseKey)) {
        New-Item -Path $baseKey -Force | Out-Null
    }
    $registryData = Get-ItemProperty -Path $baseKey -ErrorAction SilentlyContinue
    If (-Not $registryData) {
        $registryData = @{ }
    }

    $currentValues = $registryData.PSObject.Properties.Name
    If ($currentValues) {
        # Check if the extension already exists
        $existingValues = $currentValues | ForEach-Object { $registryData.$_ }
        If ($existingValues -contains $extensionString) {
            Write-Host "Extension is already installed for $browserName using ExtensionInstallForcelist." -ForegroundColor Yellow
            Return $false
        }

        # Filter only numeric keys and get the maximum value
        $numericKeys = $currentValues | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
        $newValueKey = ($numericKeys | Measure-Object -Maximum).Maximum + 1
    }
    Else {
        $newValueKey = 1
    }

    Set-ItemProperty -Path $baseKey -Name "$newValueKey" -Value $extensionString
    Write-Host "Extension installed for $browserName using ExtensionInstallForcelist." -ForegroundColor Green
}

# Remove extension from registry
Function Remove-Extension($baseKey, $browserName) {
    $removed = $false
    
    # Check if ExtensionInstallForcelist needs modification
    $forceListKey = "$baseKey\ExtensionInstallForcelist"
    If (Test-Path $forceListKey) {
        $registryData = Get-ItemProperty -Path $forceListKey -ErrorAction SilentlyContinue
        If ($registryData) {
            $currentValues = $registryData.PSObject.Properties.Name
            If ($currentValues) {
                ForEach ($value in $currentValues) {
                    $valueContent = $registryData.$value
                    If ($valueContent -like "*$global:extensionId*") {
                        Remove-ItemProperty -Path $forceListKey -Name $value
                        Write-Host "Removed ExtensionInstallForcelist key: $value for $browserName." -ForegroundColor Green
                        $removed = $true
                    }
                }
            }
        }
    }
    
    # Check if ExtensionSettings needs modification
    If (Get-ItemProperty -Path $baseKey -Name "ExtensionSettings" -ErrorAction SilentlyContinue) {
        $currentSetting = Get-ItemProperty -Path $baseKey -Name "ExtensionSettings" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ExtensionSettings
        If ($currentSetting) {
            $extSettingsJson = $currentSetting | ConvertFrom-Json
            # Check keys in ExtensionSettings
            $keys = $extSettingsJson.PSObject.Properties.Name
            $settingKeysStr = @($keys | Sort-Object) -join ","
            $targetKeysStr = @(@("*", $global:extensionId) | Sort-Object) -join ","
            If ($settingKeysStr -eq $targetKeysStr) {
                # Only "*" and the specified extension ID exist, so remove the entire ExtensionSettings value
                Remove-ItemProperty -Path $baseKey -Name "ExtensionSettings"
                Write-Host "Removed ExtensionSettings for $browserName." -ForegroundColor Green
                $removed = $true
            }
            ElseIf ($extSettingsJson.$global:extensionId) {
                # Remove just the extension ID entry
                $extSettingsJson.PSObject.Properties.Remove($global:extensionId)
                $newSettingsJson = $extSettingsJson | ConvertTo-Json -Compress
                Set-ItemProperty -Path $baseKey -Name "ExtensionSettings" -Value $newSettingsJson
                Write-Host "Removed CaseCPU extension from ExtensionSettings for $browserName." -ForegroundColor Green
                $removed = $true
            }
        }
    }
    
    If (-Not $removed) {
        Write-Host "Extension not found for $browserName." -ForegroundColor Yellow
    }
}

# Show installation status
Function Show-Status {
    Write-Host "`n=== Installation Status ===" -ForegroundColor Cyan
    
    # Check Edge
    Write-Host "`nMicrosoft Edge:" -ForegroundColor Yellow
    $edgeInstalled = $false
    
    If (Test-Path "$global:edgeKey\ExtensionInstallForcelist") {
        $registryData = Get-ItemProperty -Path "$global:edgeKey\ExtensionInstallForcelist" -ErrorAction SilentlyContinue
        If ($registryData) {
            $currentValues = $registryData.PSObject.Properties.Name
            ForEach ($value in $currentValues) {
                $valueContent = $registryData.$value
                If ($valueContent -like "*$global:extensionId*") {
                    Write-Host "  - Installed via ExtensionInstallForcelist" -ForegroundColor Green
                    $edgeInstalled = $true
                }
            }
        }
    }
    
    If (Get-ItemProperty -Path $global:edgeKey -Name "ExtensionSettings" -ErrorAction SilentlyContinue) {
        $currentSetting = Get-ItemProperty -Path $global:edgeKey -Name "ExtensionSettings" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ExtensionSettings
        If ($currentSetting) {
            $extSettingsJson = $currentSetting | ConvertFrom-Json
            If ($extSettingsJson.$global:extensionId) {
                Write-Host "  - Installed via ExtensionSettings" -ForegroundColor Green
                $edgeInstalled = $true
            }
        }
    }
    
    If (-Not $edgeInstalled) {
        Write-Host "  - Not installed" -ForegroundColor Red
    }
    
    # Check Chrome
    Write-Host "`nGoogle Chrome:" -ForegroundColor Yellow
    $chromeInstalled = $false
    
    If (Test-Path "$global:chromeKey\ExtensionInstallForcelist") {
        $registryData = Get-ItemProperty -Path "$global:chromeKey\ExtensionInstallForcelist" -ErrorAction SilentlyContinue
        If ($registryData) {
            $currentValues = $registryData.PSObject.Properties.Name
            ForEach ($value in $currentValues) {
                $valueContent = $registryData.$value
                If ($valueContent -like "*$global:extensionId*") {
                    Write-Host "  - Installed via ExtensionInstallForcelist" -ForegroundColor Green
                    $chromeInstalled = $true
                }
            }
        }
    }
    
    If (Get-ItemProperty -Path $global:chromeKey -Name "ExtensionSettings" -ErrorAction SilentlyContinue) {
        $currentSetting = Get-ItemProperty -Path $global:chromeKey -Name "ExtensionSettings" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ExtensionSettings
        If ($currentSetting) {
            $extSettingsJson = $currentSetting | ConvertFrom-Json
            If ($extSettingsJson.$global:extensionId) {
                Write-Host "  - Installed via ExtensionSettings" -ForegroundColor Green
                $chromeInstalled = $true
            }
        }
    }
    
    If (-Not $chromeInstalled) {
        Write-Host "  - Not installed" -ForegroundColor Red
    }
    
    Write-Host ""
}

# Main loop
Write-Host @"

  _____               _____ _____  _    _ 
 / ____|             / ____|  __ \| |  | |
| |     __ _ ___  ___| |    | |__) | |  | |
| |    / _` / __|/ _ \ |    |  ___/| |  | |
| |___| (_| \__ \  __/ |____| |    | |__| |
 \_____\__,_|___/\___|\_____\_|     \____/ 
                                            
Extension Installer v1.0

"@ -ForegroundColor Cyan

Write-Host "Extension ID: $global:extensionId" -ForegroundColor Gray
Write-Host ""

$exitLoop = $false
While (-Not $exitLoop) {
    $choice = Show-Menu
    Switch ($choice) {
        "1" {
            Write-Host "`nInstalling CaseCPU extension..." -ForegroundColor Cyan
            $extensionString = "$global:extensionId;https://raw.githubusercontent.com/Johnny-spec/CaseProEdgeExtension/main/updates.xml"
            
            $edgeAdded = Add-Extension1 -baseKey $global:edgeKey -browserName "Edge"
            If (-Not $edgeAdded) {
                Add-Extension2 -baseKey "$global:edgeKey\ExtensionInstallForcelist" -extensionString $extensionString -browserName "Edge"
            }

            $chromeAdded = Add-Extension1 -baseKey $global:chromeKey -browserName "Chrome"
            If (-Not $chromeAdded) {
                Add-Extension2 -baseKey "$global:chromeKey\ExtensionInstallForcelist" -extensionString $extensionString -browserName "Chrome"
            }
            
            Write-Host "`nInstallation complete! Please restart your browser." -ForegroundColor Green
        }
        "2" {
            Write-Host "`nUninstalling CaseCPU extension..." -ForegroundColor Cyan
            Remove-Extension -baseKey $global:edgeKey -browserName "Edge"
            Remove-Extension -baseKey $global:chromeKey -browserName "Chrome"
            Write-Host "`nUninstallation complete! Please restart your browser." -ForegroundColor Green
        }
        "3" {
            Show-Status
        }
        "4" {
            Write-Host "`nExiting..." -ForegroundColor Cyan
            $exitLoop = $true
        }
        Default {
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red
        }
    }
}
