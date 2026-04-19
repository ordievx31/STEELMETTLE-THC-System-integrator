# =============================================================================
# MACH3/MACH4 STEELMETTLE THC Feedrate Macro Installer
# =============================================================================
# This module provides functions to detect Mach3/Mach4 installations and
# automatically install the STEELMETTLE THC feedrate integration macros
#
# Usage:
#   $result = Install-MachFeedrateMacro -MachVersion "Mach3" -ProfileName (optional)
#   $result = Install-MachFeedrateMacro -MachVersion "Mach4"
#
# =============================================================================

function Find-MachInstallation([string]$MachVersion = "Mach3") {
    <#
    .SYNOPSIS
    Locate Mach3 or Mach4 installation directory
    
    .PARAMETER MachVersion
    "Mach3" or "Mach4" - specify which CNC software to search for
    
    .OUTPUTS
    @{ Path = "C:\Mach3"; Type = "Mach3"; Description = "..." } or $null
    #>
    
    $candidates = @()
    
    if ($MachVersion -eq "Mach3") {
        $candidates = @(
            "C:\Mach3",
            "C:\Mach3Dev",
            "$env:ProgramFiles\Mach3",
            "$env:ProgramFiles\Mach3Dev"
        )
    } elseif ($MachVersion -eq "Mach4") {
        $candidates = @(
            "C:\Mach4",
            "C:\Mach4Hobby",
            "$env:ProgramFiles\Mach4",
            "$env:ProgramFiles\Mach4Hobby"
        )
    }
    
    foreach ($path in $candidates) {
        if (Test-Path $path) {
            $exe = if ($MachVersion -eq "Mach3") { "Mach3.exe" } else { "Mach4.exe" }
            if (Test-Path (Join-Path $path $exe)) {
                return @{
                    Path = $path
                    Type = $MachVersion
                    Description = "$MachVersion installation found at $path"
                }
            }
        }
    }
    
    return $null
}

function Find-MachMacrosFolder([string]$MachInstallPath, [string]$MachVersion) {
    <#
    .SYNOPSIS
    Locate Macros folder within Mach3/Mach4 installation
    For Mach3: typically [MachPath]\Macros
    For Mach4: typically [MachPath]\Profiles\[ProfileName]\Macros
    #>
    
    if ($MachVersion -eq "Mach3") {
        # Mach3 macros folder is at root level
        $macrosPath = Join-Path $MachInstallPath "Macros"
        if (Test-Path $macrosPath) {
            return $macrosPath
        }
    } elseif ($MachVersion -eq "Mach4") {
        # Mach4 macros are in Profiles subdirectories
        $profilesPath = Join-Path $MachInstallPath "Profiles"
        if (Test-Path $profilesPath) {
            # List available profiles
            $profiles = Get-ChildItem $profilesPath -Directory | Select-Object -ExpandProperty Name
            if ($profiles.Count -gt 0) {
                # Default to first profile or "Default"
                $defaultProfile = if ($profiles -contains "Default") { "Default" } else { $profiles[0] }
                return (Join-Path $profilesPath $defaultProfile "Macros"), $defaultProfile
            }
        }
    }
    
    return $null
}

function New-MachMacroInstallationDialog([string]$MachVersion, [array]$AvailableInstallations) {
    <#
    .SYNOPSIS
    Display WinForms dialog to select which Mach installation to target
    #>
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "STEELMETTLE THC: Install $MachVersion Macro"
    $form.Width = 500
    $form.Height = 280
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    
    # Style colors
    $accentColor = [System.Drawing.Color]::FromArgb(88, 84, 40)
    $textColor = [System.Drawing.Color]::White
    
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Select $MachVersion Installation:"
    $lbl.Location = [System.Drawing.Point]::new(16, 18)
    $lbl.Size = [System.Drawing.Size]::new(450, 24)
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lbl)
    
    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.Location = [System.Drawing.Point]::new(16, 50)
    $cmb.Size = [System.Drawing.Size]::new(450, 30)
    $cmb.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    
    foreach ($inst in $AvailableInstallations) {
        [void]$cmb.Items.Add($inst.Description)
    }
    
    if ($cmb.Items.Count -gt 0) {
        $cmb.SelectedIndex = 0
    }
    $form.Controls.Add($cmb)
    
    $instrLbl = New-Object System.Windows.Forms.Label
    $instrLbl.Text = @"
The macro file will be copied and startup
configuration will be applied automatically.

Restart $MachVersion after installation to activate.
"@
    $instrLbl.Location = [System.Drawing.Point]::new(16, 92)
    $instrLbl.Size = [System.Drawing.Size]::new(450, 100)
    $instrLbl.TextAlign = "TopLeft"
    $form.Controls.Add($instrLbl)
    
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Install"
    $btnOk.Location = [System.Drawing.Point]::new(278, 204)
    $btnOk.Size = [System.Drawing.Size]::new(90, 32)
    $btnOk.BackColor = $accentColor
    $btnOk.ForeColor = $textColor
    $btnOk.FlatStyle = "Flat"
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $btnOk
    $form.Controls.Add($btnOk)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = [System.Drawing.Point]::new(376, 204)
    $btnCancel.Size = [System.Drawing.Size]::new(90, 32)
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(122, 58, 34)
    $btnCancel.ForeColor = $textColor
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancel
    $form.Controls.Add($btnCancel)
    
    $result = $form.ShowDialog()
    $selectedIndex = $cmb.SelectedIndex
    $form.Dispose()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $selectedIndex -ge 0) {
        return $AvailableInstallations[$selectedIndex]
    }
    
    return $null
}

function Configure-Mach3Startup([string]$MachInstallPath) {
    <#
    .SYNOPSIS
    Inject SetTimer startup/shutdown lines into Mach3 profile XML files.
    Only targets files that contain a <Profile> root element (real Mach3 profiles).
    Creates a .bak backup before modifying. Idempotent - safe to run multiple times.
    .OUTPUTS
    @{ Configured = $true/$false; Profiles = @("MILL.xml",...); Message = "..." }
    #>
    $startupLine   = 'SetTimer 1, 50, SendParametersToTHC'
    $shutdownLine  = 'SetTimer 1, 0'
    $startupEsc    = [regex]::Escape($startupLine)
    $shutdownEsc   = [regex]::Escape($shutdownLine)
    $configured    = @()
    $errors        = @()

    # Get all root-level XML files, excluding known non-profile files
    $xmlFiles = Get-ChildItem -Path $MachInstallPath -Filter '*.xml' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.DirectoryName -eq $MachInstallPath -and $_.Name -ne 'uninstall.xml' }

    if (-not $xmlFiles -or $xmlFiles.Count -eq 0) {
        return @{ Configured = $false; Profiles = @(); Message = "No XML files found in $MachInstallPath" }
    }

    foreach ($pf in $xmlFiles) {
        try {
            $content = [System.IO.File]::ReadAllText($pf.FullName, [System.Text.Encoding]::UTF8)

            # BUG 1 FIX: Only process files that are actual Mach3 profile XMLs
            if ($content -notmatch '<Profile\b') {
                continue
            }

            # Idempotency: skip if both lines are already present
            $hasStartup  = $content -match $startupEsc
            $hasShutdown = $content -match $shutdownEsc
            if ($hasStartup -and $hasShutdown) {
                $configured += $pf.Name
                continue
            }

            # Create backup before modifying
            $bakPath = $pf.FullName + '.bak'
            if (-not (Test-Path $bakPath)) {
                [System.IO.File]::Copy($pf.FullName, $bakPath)
            }

            $changed = $false

            # --- Startup injection (single pass) ---
            if (-not $hasStartup) {
                # Detect existing startup tag (multiline-safe with (?s))
                $startTag = $null
                foreach ($tag in @('StartupCode','StartScript','OnStartScript','MacroStartup')) {
                    if ($content -match "<$tag\b") { $startTag = $tag; break }
                }

                if ($startTag) {
                    # Tag exists - append our line inside it ((?s) so .*? spans newlines)
                    $content = $content -replace "(?s)(<$startTag>)(.*?)(</$startTag>)", "`$1`$2`r`n$startupLine`r`n`$3"
                    $changed = $true
                } elseif ($content -match '</[Pp]rofile>') {
                    # No startup tag - create one before closing </Profile>
                    $content = $content -replace '(</[Pp]rofile>)', "  <StartupCode>$startupLine</StartupCode>`r`n`$1"
                    $changed = $true
                }
            }

            # --- Shutdown injection (single pass, independent of startup) ---
            if (-not $hasShutdown) {
                $shutTag = $null
                foreach ($tag in @('ShutdownCode','ShutdownScript','OnShutdownScript','MacroShutdown')) {
                    if ($content -match "<$tag\b") { $shutTag = $tag; break }
                }

                if ($shutTag) {
                    $content = $content -replace "(?s)(<$shutTag>)(.*?)(</$shutTag>)", "`$1`$2`r`n$shutdownLine`r`n`$3"
                    $changed = $true
                } elseif ($content -match '</[Pp]rofile>') {
                    $content = $content -replace '(</[Pp]rofile>)', "  <ShutdownCode>$shutdownLine</ShutdownCode>`r`n`$1"
                    $changed = $true
                }
            }

            if ($changed) {
                [System.IO.File]::WriteAllText($pf.FullName, $content, [System.Text.Encoding]::UTF8)
            }
            $configured += $pf.Name
        } catch {
            $errors += "$($pf.Name): $_"
        }
    }

    if ($configured.Count -eq 0 -and $errors.Count -eq 0) {
        return @{ Configured = $false; Profiles = @(); Message = "No Mach3 profile XML files (containing <Profile>) found in $MachInstallPath" }
    }
    if ($errors.Count -gt 0) {
        return @{ Configured = ($configured.Count -gt 0); Profiles = $configured; Message = "Errors: $($errors -join '; ')" }
    }
    return @{ Configured = $true; Profiles = $configured; Message = "Startup timer configured in: $($configured -join ', ')" }
}

function Configure-Mach4Startup([string]$MacrosFolder) {
    <#
    .SYNOPSIS
    Inject dofile() and PLC call into Mach4 profile scripts
    .OUTPUTS
    @{ Configured = $true/$false; Message = "..." }
    #>
    $dofileLine = 'dofile(inst:GetMachDir() .. "\\Profiles\\" .. inst:GetProfileName() .. "\\Macros\\STEELMETTLE_THC_Feedrate.lua")'
    $callLine   = '  SendParametersToTHC()'
    $marker     = 'STEELMETTLE_THC_Feedrate'

    # Target: mcPLC.lua in the same Macros folder
    $plcFile = Join-Path $MacrosFolder "mcPLC.lua"
    $configured = $false

    try {
        if (Test-Path $plcFile) {
            $content = [System.IO.File]::ReadAllText($plcFile, [System.Text.Encoding]::UTF8)
            if ($content -notmatch [regex]::Escape($marker)) {
                # Insert dofile at top and call inside PLC function
                $header = "-- STEELMETTLE THC Feedrate Integration (auto-installed)`r`n$dofileLine`r`n"
                $content = $header + $content
                # Try to inject call inside function plc_run() or similar
                if ($content -match '(function\s+[Pp][Ll][Cc]_?[Rr]un\s*\([^)]*\)[^\r\n]*)') {
                    $content = $content -replace '(function\s+[Pp][Ll][Cc]_?[Rr]un\s*\([^)]*\)[^\r\n]*)', "`$1`r`n$callLine"
                }
                [System.IO.File]::WriteAllText($plcFile, $content, [System.Text.Encoding]::UTF8)
                $configured = $true
            } else {
                $configured = $true  # Already present
            }
        } else {
            # Create a minimal mcPLC.lua with our integration
            $newContent = @"
-- STEELMETTLE THC Feedrate Integration (auto-installed)
$dofileLine

function mcPLC_run()
$callLine
end
"@
            [System.IO.File]::WriteAllText($plcFile, $newContent, [System.Text.Encoding]::UTF8)
            $configured = $true
        }
    } catch {
        return @{ Configured = $false; Message = "Failed to configure mcPLC.lua: $_" }
    }

    return @{ Configured = $true; Message = "Mach4 PLC script configured for auto-start" }
}

function Show-MacroInstallationInstructions([string]$MachVersion, [string]$MacrosFolder, [string]$MacroFileName, [string]$MacroDisplayPath, [bool]$AutoConfigured = $false, [string]$AutoConfigMessage = "") {
    <#
    .SYNOPSIS
    Display a message box with installation result
    #>
    
    Add-Type -AssemblyName System.Windows.Forms
    
    if ($AutoConfigured) {
        $instructions = @"
MACRO INSTALLATION COMPLETE

Macro file copied to:
  $MacroDisplayPath

Startup auto-configured:
  $AutoConfigMessage

Everything is set up. Restart $MachVersion to activate
the STEELMETTLE THC feedrate integration.
"@
    } else {
        $instructions = @"
MACRO FILE INSTALLED

Macro file copied to:
  $MacroDisplayPath

Auto-configuration was not possible:
  $AutoConfigMessage

$( if ($MachVersion -eq "Mach3") {
@"
MANUAL STEP REQUIRED:
  1. Open Mach3
  2. Go to Config -> System Startup/Shutdown Macros
  3. Add to Startup:   SetTimer 1, 50, SendParametersToTHC
  4. Add to Shutdown:  SetTimer 1, 0
  5. Restart Mach3
"@
} else {
@"
MANUAL STEP REQUIRED:
  1. Open your Mach4 profile Macros folder
  2. Edit mcPLC.lua
  3. Add at top:  dofile(inst:GetMachDir() .. "\\Profiles\\" .. inst:GetProfileName() .. "\\Macros\\STEELMETTLE_THC_Feedrate.lua")
  4. Call SendParametersToTHC() in plc_run()
  5. Restart Mach4
"@
})
"@
    }
    
    $title = if ($AutoConfigured) { "STEELMETTLE THC: $MachVersion Fully Configured" } else { "STEELMETTLE THC: Manual Step Required" }
    $icon  = if ($AutoConfigured) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
    
    [System.Windows.Forms.MessageBox]::Show(
        $instructions,
        $title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $icon
    ) | Out-Null
}

function Install-MachFeedrateMacro {
    <#
    .SYNOPSIS
    Main function: Detect Mach installation, copy macro file, show instructions
    
    .PARAMETER MachVersion
    "Mach3" or "Mach4"
    
    .OUTPUTS
    @{ Success = $true; Message = "..."; MacroPath = "..." } or failure object
    #>
    
    param(
        [ValidateSet("Mach3", "Mach4")]
        [string]$MachVersion = "Mach3"
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    
    # Step 1: Locate Mach installation
    $installation = Find-MachInstallation -MachVersion $MachVersion
    
    if (-not $installation) {
        [System.Windows.Forms.MessageBox]::Show(
            "$MachVersion installation not found.`n`nSearched common locations:`n- C:\$MachVersion`n- C:\Program Files\$MachVersion`n`nPlease install $MachVersion or check the path.",
            "STEELMETTLE THC: $MachVersion Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        
        return @{ Success = $false; Message = "$MachVersion not found" }
    }
    
    # Step 2: Confirm installation with user (in case multiple exist)
    # For now, use the detected one; could be enhanced to show dialog if multiple found
    
    # Step 3: Locate Macros folder
    $macrosFolderResult = Find-MachMacrosFolder -MachInstallPath $installation.Path -MachVersion $MachVersion
    
    if (-not $macrosFolderResult) {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not locate the Macros folder in $MachVersion installation.`n`nPath searched: $($installation.Path)`n`nPlease check your $MachVersion installation.",
            "STEELMETTLE THC: Macros Folder Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        
        return @{ Success = $false; Message = "Macros folder not found" }
    }
    
    # macrosFolderResult could be string or array [path, profileName]
    $macrosFolder = if ($macrosFolderResult -is [array]) { $macrosFolderResult[0] } else { $macrosFolderResult }
    
    # Create Macros folder if it doesn't exist
    if (-not (Test-Path $macrosFolder)) {
        try {
            New-Item -ItemType Directory -Path $macrosFolder -Force | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to create Macros folder:`n`n$_",
                "STEELMETTLE THC: Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            
            return @{ Success = $false; Message = "Failed to create Macros folder" }
        }
    }
    
    # Step 4: Determine source and destination file names
    $sourceFileName = if ($MachVersion -eq "Mach3") { "MACH3_FEEDRATE_THC_MACRO.m3s" } else { "MACH4_FEEDRATE_THC_MACRO.lua" }
    $destFileName = if ($MachVersion -eq "Mach3") { "STEELMETTLE_THC_Feedrate.m3s" } else { "STEELMETTLE_THC_Feedrate.lua" }
    
    # Step 5: Find source template in integrator
    # $baseDir is defined by the main script; fall back to $PSScriptRoot if available
    $integrationDir = if ($baseDir) { $baseDir } elseif ($PSScriptRoot) { Split-Path -Parent $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $sourceFile = Join-Path (Join-Path $integrationDir "PoKeys") $sourceFileName
    
    if (-not (Test-Path $sourceFile)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Macro template not found:`n`n$sourceFile`n`nThe STEELMETTLE integrator may be incomplete.",
            "STEELMETTLE THC: Template Missing",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        
        return @{ Success = $false; Message = "Macro template not found" }
    }
    
    # Step 6: Copy macro file
    $destFile = Join-Path $macrosFolder $destFileName
    
    try {
        Copy-Item -Path $sourceFile -Destination $destFile -Force
        
        # Write log entry
        $logMsg = "Copied $MachVersion macro: $sourceFile -> $destFile"
        if ($null -ne $logDir) {
            (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + " $logMsg" | Add-Content (Join-Path $logDir 'integrator.log') -Encoding ASCII
        }
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to copy macro file:`n`n$_`n`nCheck that $macrosFolder is writable.",
            "STEELMETTLE THC: Copy Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        
        return @{ Success = $false; Message = "Failed to copy macro file: $_" }
    }
    
    # Step 7: Auto-configure startup integration
    $autoConfigured = $false
    $autoConfigMsg  = ""

    if ($MachVersion -eq "Mach3") {
        $cfgResult = Configure-Mach3Startup -MachInstallPath $installation.Path
        $autoConfigured = $cfgResult.Configured
        $autoConfigMsg  = $cfgResult.Message
    } else {
        $cfgResult = Configure-Mach4Startup -MacrosFolder $macrosFolder
        $autoConfigured = $cfgResult.Configured
        $autoConfigMsg  = $cfgResult.Message
    }

    # Step 8: Show result
    $displayPath = $destFile
    if ($macrosFolderResult -is [array]) {
        $displayPath = "...\Profiles\$($macrosFolderResult[1])\Macros\$destFileName"
    }
    
    Show-MacroInstallationInstructions -MachVersion $MachVersion -MacrosFolder $macrosFolder -MacroFileName $destFileName -MacroDisplayPath $displayPath -AutoConfigured $autoConfigured -AutoConfigMessage $autoConfigMsg
    
    return @{
        Success = $true
        Message = "Macro installed successfully"
        MacroPath = $destFile
        MachPath = $installation.Path
        MacrosFolder = $macrosFolder
        AutoConfigured = $autoConfigured
    }
}

if ($null -ne $ExecutionContext.SessionState.Module) {
    Export-ModuleMember -Function @(
        'Find-MachInstallation',
        'Find-MachMacrosFolder',
        'Install-MachFeedrateMacro'
    )
}

