<#
PowerShell integrator script for STM32/Arduino/PoKeys firmware management.

Run this from the folder containing this script:
  .\run-integrator.ps1

It will attempt to detect connected devices and run the appropriate build/flash step.

Requirements (copy to this folder):
  - STM32_Programmer_CLI.exe (for STM32 flashing)
  - arduino-cli.exe (for Arduino build/upload)
  - Optional: pokeys-cli.exe (or equivalent) to load PoKeys macros.
#>

param(
    [switch]$Auto,    # If set, run automatically without prompting (one-click)
    [switch]$Gui      # If set, show a simple GUI to choose actions
)

$ErrorActionPreference = 'Stop'

# When run as an embedded script string (inside the compiled EXE), $MyInvocation.MyCommand.Path
# is null. The C# launcher injects $global:__SMTExeBaseDir for exactly this case.
$baseDir = if ($global:__SMTExeBaseDir) {
    $global:__SMTExeBaseDir
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $PWD.Path
}
$toolsDir = Join-Path $baseDir 'tools'
$logDir = Join-Path $baseDir 'logs'
$docsDir = Split-Path -Parent $baseDir

# Mach macro installer helper path (loaded eagerly at script scope)
$machMacroModulePath = Join-Path $baseDir 'MachMacroInstaller.ps1'
$script:machMacroModuleLoaded = $false
if (Test-Path $machMacroModulePath) {
    try {
        . $machMacroModulePath
        $script:machMacroModuleLoaded = $true
    } catch {
        try {
            $machMacroRaw = Get-Content -Path $machMacroModulePath -Raw -ErrorAction Stop
            . ([scriptblock]::Create($machMacroRaw))
            $script:machMacroModuleLoaded = $true
        } catch {
            Write-Host "WARNING: Failed to load Mach macro helper: $_" -ForegroundColor Yellow
        }
    }
}

function Ensure-MachMacroModuleLoaded {
    return $script:machMacroModuleLoaded
}

function ConvertTo-CompatiblePsObject($value) {
    if ($null -eq $value) { return $null }

    if ($value -is [System.Collections.IDictionary]) {
        $ht = [ordered]@{}
        foreach ($k in $value.Keys) {
            $ht[[string]$k] = ConvertTo-CompatiblePsObject $value[$k]
        }
        return [pscustomobject]$ht
    }

    if (($value -is [System.Collections.IEnumerable]) -and -not ($value -is [string])) {
        $arr = @()
        foreach ($item in $value) {
            $arr += ,(ConvertTo-CompatiblePsObject $item)
        }
        return $arr
    }

    return $value
}

function ConvertFrom-JsonCompat([string]$jsonText) {
    if ([string]::IsNullOrWhiteSpace($jsonText)) { return $null }

    try {
        return ConvertFrom-Json -InputObject $jsonText -ErrorAction Stop
    } catch {
        # Compatibility fallback for environments where ConvertFrom-Json is missing/limited.
        Add-Type -AssemblyName System.Web.Extensions -ErrorAction Stop
        $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $parsed = $ser.DeserializeObject($jsonText)
        return ConvertTo-CompatiblePsObject $parsed
    }
}

# Optional config file for overrides (paths, FQBN, etc.)
$configPath = Join-Path $baseDir 'config.json'
$config = $null
if (Test-Path $configPath) {
    try {
        $config = ConvertFrom-JsonCompat (Get-Content $configPath -Raw)
    } catch {
        Write-Host "WARNING: Failed to parse config.json: $_" -ForegroundColor Yellow
    }
}

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

$script:devModeUnlocked = $false
$script:decodeModeUnlocked = $false
$script:pendingLockTargets = @{}
$script:isAppUpdateShutdown = $false
$script:devLicenseImportedThisRun = $false

function Restart-IntegratorApp {
    $launched = $false

    try {
        $currentExe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if (-not [string]::IsNullOrWhiteSpace($currentExe) -and (Test-Path $currentExe)) {
            $leaf = [System.IO.Path]::GetFileName($currentExe)
            if ($leaf -like 'STEELMETTLE-THC-Systems-Integrator*.exe') {
                Start-Process -FilePath $currentExe | Out-Null
                $launched = $true
            }
        }
    } catch {}

    if (-not $launched -and $MyInvocation.MyCommand.Path -and (Test-Path $MyInvocation.MyCommand.Path)) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $MyInvocation.MyCommand.Path) | Out-Null
        $launched = $true
    }

    if (-not $launched) {
        $fallbackExe = Join-Path $baseDir 'STEELMETTLE-THC-Systems-Integrator.exe'
        if (Test-Path $fallbackExe) {
            Start-Process -FilePath $fallbackExe | Out-Null
            $launched = $true
        }
    }

    return $launched
}

function Get-TransferProtectionSeed {
    return 'STEELMETTLE|THC|Systems|Integrator|Transfer|Protection|2026|v1'
}

function Get-TransferKeyBytes([string]$scope) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $material = [System.Text.Encoding]::UTF8.GetBytes((Get-TransferProtectionSeed) + '|' + $scope)
        return $sha.ComputeHash($material)
    } finally {
        $sha.Dispose()
    }
}

function Unprotect-TransferBytes([byte[]]$payloadBytes, [string]$scope) {
    if (-not $payloadBytes -or $payloadBytes.Length -le 16) {
        throw 'Encrypted payload is invalid.'
    }

    $aes = [System.Security.Cryptography.Aes]::Create()
    try {
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        $aes.KeySize = 256
        $aes.BlockSize = 128
        $aes.Key = Get-TransferKeyBytes $scope

        $iv = New-Object byte[] 16
        [System.Array]::Copy($payloadBytes, 0, $iv, 0, 16)
        $aes.IV = $iv

        $cipherBytes = New-Object byte[] ($payloadBytes.Length - 16)
        [System.Array]::Copy($payloadBytes, 16, $cipherBytes, 0, $cipherBytes.Length)

        $mem = New-Object System.IO.MemoryStream
        try {
            $dec = $aes.CreateDecryptor()
            try {
                $crypto = New-Object System.Security.Cryptography.CryptoStream($mem, $dec, [System.Security.Cryptography.CryptoStreamMode]::Write)
                try {
                    $crypto.Write($cipherBytes, 0, $cipherBytes.Length)
                    $crypto.FlushFinalBlock()
                } finally {
                    $crypto.Dispose()
                }
            } finally {
                $dec.Dispose()
            }

            return $mem.ToArray()
        } finally {
            $mem.Dispose()
        }
    } finally {
        $aes.Dispose()
    }
}

function Decode-TransferPayload([string]$raw, [string]$prefix, [string]$scope) {
    if (-not $raw) { return $null }
    $trimmed = $raw.Trim()
    if (-not $trimmed.StartsWith($prefix)) { return $null }

    $payloadBytes = [System.Convert]::FromBase64String($trimmed.Substring($prefix.Length))
    try {
        return Unprotect-TransferBytes $payloadBytes $scope
    } catch {
        return $payloadBytes
    }
}

function Mark-TargetPendingLock([ValidateSet('Forge','Core')][string]$Target) {
    $script:pendingLockTargets[$Target] = $true
}

function Clear-TargetPendingLock([ValidateSet('Forge','Core')][string]$Target) {
    if ($script:pendingLockTargets.ContainsKey($Target)) {
        $script:pendingLockTargets.Remove($Target) | Out-Null
    }
}

function Get-PendingLockTargets {
    return @($script:pendingLockTargets.Keys | Sort-Object)
}

function Test-PendingDevLocksRequireAction {
    if (-not $script:devModeUnlocked) { return $false }

    $sec = Get-SecurityConfig
    if ($null -ne $sec.enforcePendingLocksOnExit -and -not $sec.enforcePendingLocksOnExit) {
        return $false
    }

    $tier = Get-LicenseTier
    # User tier never has pending locks (no dev flash capability)
    if ($tier -eq 'user') { return $false }

    return ((Get-PendingLockTargets).Count -gt 0)
}

function Ensure-PendingDevLocksBeforeExit {
    param(
        [System.Windows.Forms.Form]$OwnerForm
    )

    $targets = Get-PendingLockTargets
    if ($targets.Count -eq 0) {
        return $true
    }

    $sec = Get-SecurityConfig
    if (-not $sec.enabled) {
        [System.Windows.Forms.MessageBox]::Show(
            "Developer mode flashed unlocked firmware for: $($targets -join ', ').`n`nProduction lock enforcement is disabled in config.security.enabled, so the app cannot apply the required lock before exit.",
            'Pending Production Lock',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $true
    }

    # Owner tier: remind but allow skip. Developer tier: must lock.
    $tier = Get-LicenseTier

    if ($tier -eq 'owner') {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Unlocked firmware was flashed for: $($targets -join ', ').`n`nIs this a production unit that needs to be locked?`n`n  Yes = Apply production lock now`n  No = Skip lock and close`n  Cancel = Stay in the app",
            'Production Lock Reminder',
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
            return $false
        }
        if ($result -eq [System.Windows.Forms.DialogResult]::No) {
            Log "Owner skipped production lock for: $($targets -join ', ')"
            return $true
        }
        # Yes - fall through to lock logic below
    } else {
        # Developer tier: must lock before exit
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Developer mode flashed unlocked firmware for: $($targets -join ', ').`n`nThe device must be production-locked before exiting. Click OK to apply the required lock(s), or Cancel to stay in the app.",
            'Pending Production Lock',
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            return $false
        }
    }

    foreach ($target in @($targets)) {
        try {
            Set-UiStage "Applying $target production lock before exit..." 90
            Apply-ProductionLock -Target $target
            Clear-TargetPendingLock -Target $target
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to apply $target production lock before exit:`n$_",
                'Pending Production Lock',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            return $false
        }
    }

    Set-UiStage 'Required production lock(s) applied.' 100
    return $true
}

function Log($text) {
    $time = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$time] $text"
    $line | Tee-Object -FilePath (Join-Path $logDir 'integrator.log') -Append | Out-Null
}

$script:UiProgressCallback = $null

function Set-UiStage([string]$text, [int]$value = -1) {
    if ($script:UiProgressCallback) {
        & $script:UiProgressCallback $text $value
    }
}

function Get-UiConfig {
    function Resolve-UiPath([string]$p) {
        if ([string]::IsNullOrWhiteSpace($p)) { return $null }
        if ([System.IO.Path]::IsPathRooted($p)) { return $p }
        return (Join-Path $baseDir $p)
    }

    $splashImage = if ($config -and $config.ui -and $config.ui.splashImage) {
        Resolve-UiPath $config.ui.splashImage
    } else {
        Resolve-UiPath '..\OneDrive\Pictures\SteelMettleLLC_copy1.jpg'
    }

    $appIcon = if ($config -and $config.ui -and $config.ui.appIcon) {
        Resolve-UiPath $config.ui.appIcon
    } else {
        Resolve-UiPath 'assets\SteelMettle.ico'
    }

    $forgeScreenImage = if ($config -and $config.ui -and $config.ui.forgeScreenImage) {
        Resolve-UiPath $config.ui.forgeScreenImage
    } else {
        'C:\Users\jashu\OneDrive\Pictures\SteelMettleLLC-20_copy1_copy.jpg'
    }

    return @{
        splashImage = $splashImage
        appIcon = $appIcon
        forgeScreenImage = $forgeScreenImage
    }
}

function Get-SyncConfig {
    function Resolve-SyncPath([string]$p) {
        if ([string]::IsNullOrWhiteSpace($p)) { return $p }
        if ([System.IO.Path]::IsPathRooted($p)) { return $p }
        return (Join-Path $baseDir $p)
    }

    $syncEnabled = $true
    if ($config -and $config.sync -and $null -ne $config.sync.enabled) {
        $syncEnabled = [bool]$config.sync.enabled
    }

    $arduinoSourceDir = if ($config -and $config.sync -and $config.sync.arduinoSourceDir) {
        Resolve-SyncPath $config.sync.arduinoSourceDir
    } else {
        Join-Path $docsDir 'Arduino\THC_AUTOSET_Firmware'
    }

    $stm32SourceBin = if ($config -and $config.sync -and $config.sync.stm32SourceBin) {
        Resolve-SyncPath $config.sync.stm32SourceBin
    } else {
        Join-Path $docsDir 'STM32MX_V3.0.1\Debug\STM32MX_V2.bin'
    }

    $stm32SourceConfigHeader = if ($config -and $config.sync -and $config.sync.stm32SourceConfigHeader) {
        Resolve-SyncPath $config.sync.stm32SourceConfigHeader
    } else {
        Join-Path $docsDir 'STM32MX_V3.0.1\Core\Inc\THC_Config.h'
    }

    $pokeysMacroSourceFile = if ($config -and $config.sync -and $config.sync.pokeysMacroSourceFile) {
        Resolve-SyncPath $config.sync.pokeysMacroSourceFile
    } else {
        Join-Path $docsDir 'Arduino\THC_AUTOSET_Firmware\POKEYS_POEXT_MACROS_CLEAN.txt'
    }

    return @{
        enabled = $syncEnabled
        arduinoSourceDir = $arduinoSourceDir
        stm32SourceBin = $stm32SourceBin
        stm32SourceConfigHeader = $stm32SourceConfigHeader
        pokeysMacroSourceFile = $pokeysMacroSourceFile
    }
}

function Get-FileSyncHash {
    param([string]$Path)
    try {
        if (Test-Path $Path -PathType Leaf) {
            return (Get-FileHash -Algorithm SHA256 -Path $Path -ErrorAction Stop).Hash
        }
    } catch {
        return $null
    }
    return $null
}

function Load-SyncHashes {
    param([string]$CacheFile)
    $hashes = @{}
    if (-not (Test-Path $CacheFile)) {
        return $hashes
    }

    try {
        $json = Get-Content $CacheFile -Raw -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($json)) {
            $obj = ConvertFrom-JsonCompat $json
            if ($obj) {
                foreach ($p in $obj.PSObject.Properties) {
                    $hashes[[string]$p.Name] = [string]$p.Value
                }
            }
        }
    } catch {
        Log "WARNING: Could not load sync hash cache, starting fresh: $_"
    }

    return $hashes
}

function Save-SyncHashes {
    param(
        [hashtable]$Hashes,
        [string]$CacheFile
    )

    $dir = Split-Path -Parent $CacheFile
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    try {
        $out = [ordered]@{}
        foreach ($k in ($Hashes.Keys | Sort-Object)) {
            $out[$k] = $Hashes[$k]
        }
        $out | ConvertTo-Json -Depth 5 | Set-Content -Path $CacheFile -ErrorAction Stop
    } catch {
        Log "WARNING: Could not save sync hash cache: $_"
    }
}

function Get-Stm32AiOptions {
    $ai = if ($config -and $config.stm32 -and $config.stm32.ai) { $config.stm32.ai } else { $null }
    return @{
        enabled = if ($ai -and $null -ne $ai.enabled) { [bool]$ai.enabled } else { $true }
        predictEnable = if ($ai -and $null -ne $ai.predictEnable) { [bool]$ai.predictEnable } else { $true }
        predictWindowSize = if ($ai -and $null -ne $ai.predictWindowSize) { [int]$ai.predictWindowSize } else { 10 }
        learningRate = if ($ai -and $null -ne $ai.learningRate) { [double]$ai.learningRate } else { 0.1 }
        thresholdV = if ($ai -and $null -ne $ai.thresholdV) { [double]$ai.thresholdV } else { 5.0 }
        confidenceMin = if ($ai -and $null -ne $ai.confidenceMin) { [double]$ai.confidenceMin } else { 0.7 }
    }
}

function Get-UpdateConfig {
    $u = if ($config -and $config.update) { $config.update } else { $null }
    return @{
        enabled = if ($u -and $null -ne $u.enabled) { [bool]$u.enabled } else { $false }
        currentVersion = if ($u -and $u.currentVersion) { [string]$u.currentVersion } else { '0.0.0' }
        repo = if ($u -and $u.repo) { [string]$u.repo } else { '' }
        assetNamePattern = if ($u -and $u.assetNamePattern) { [string]$u.assetNamePattern } else { '' }
        autoCheckOnLaunch = if ($u -and $null -ne $u.autoCheckOnLaunch) { [bool]$u.autoCheckOnLaunch } else { $false }
    }
}

function Get-SecurityConfig {
    $s = if ($config -and $config.security) { $config.security } else { $null }
    return @{
        enabled                   = if ($s -and $null -ne $s.enabled) { [bool]$s.enabled } else { $false }
        requireConfirmation       = if ($s -and $null -ne $s.requireConfirmation) { [bool]$s.requireConfirmation } else { $true }
        enforcePendingLocksOnExit = if ($s -and $null -ne $s.enforcePendingLocksOnExit) { [bool]$s.enforcePendingLocksOnExit } else { $true }
        commandTimeoutSeconds     = if ($s -and $null -ne $s.commandTimeoutSeconds) { [int]$s.commandTimeoutSeconds } else { 45 }
        userLicenseFile           = if ($s -and $s.userLicenseFile) { [string]$s.userLicenseFile } else { 'licenses/user.license.json' }
        devLicenseFile            = if ($s -and $s.devLicenseFile) { [string]$s.devLicenseFile } else { 'licenses/dev.license.json' }
        forgeLockCommand          = if ($s -and $s.forgeLockCommand) { [string]$s.forgeLockCommand } else { '' }
        coreLockCommand           = if ($s -and $s.coreLockCommand) { [string]$s.coreLockCommand } else { '' }
    }
}

function Resolve-SecurityPath([string]$pathValue, [string]$defaultRelativePath) {
    $effective = if ([string]::IsNullOrWhiteSpace($pathValue)) { $defaultRelativePath } else { $pathValue }
    if ([System.IO.Path]::IsPathRooted($effective)) {
        return $effective
    }
    return (Join-Path $baseDir $effective)
}

function Get-UserLicensePath {
    $sec = Get-SecurityConfig
    return (Resolve-SecurityPath -pathValue $sec.userLicenseFile -defaultRelativePath 'licenses/user.license.json')
}

function Get-DevLicensePath {
    $sec = Get-SecurityConfig
    return (Resolve-SecurityPath -pathValue $sec.devLicenseFile -defaultRelativePath 'licenses/dev.license.json')
}

function Read-LicenseFile([string]$licensePath) {
    if (-not (Test-Path $licensePath)) {
        return $null
    }

    try {
        $raw = Get-Content -Path $licensePath -Raw
        if ($raw -and $raw.TrimStart().StartsWith('SMTLIC:')) {
            $bytes = Decode-TransferPayload -raw $raw -prefix 'SMTLIC:' -scope 'license'
            $raw = [System.Text.Encoding]::UTF8.GetString($bytes)
        }
        # Some encoded/decoded payloads may retain a BOM/control prefix; normalize before JSON parse.
        $raw = ([string]$raw).TrimStart([char]0xFEFF, [char]0x200B, [char]0x2060, [char]0)
        return (ConvertFrom-JsonCompat $raw)
    } catch {
        throw "License file parse failed: $licensePath"
    }
}

function Ensure-UserLicensePresent {
    $path = Get-UserLicensePath
    $obj = Read-LicenseFile $path
    if (-not $obj) {
        throw "User license file is required and was not found: $path"
    }

    if (-not $obj.licenseType -or ([string]$obj.licenseType).ToLower() -ne 'user') {
        throw "User license file is invalid (licenseType must be 'user'): $path"
    }
}

function Get-DevLicenseHash {
    $path = Get-DevLicensePath
    $obj = Read-LicenseFile $path
    if (-not $obj) { return '' }
    if (-not $obj.licenseType -or ([string]$obj.licenseType).ToLower() -ne 'developer') { return '' }
    if (-not $obj.devKeyHash) { return '' }
    return [string]$obj.devKeyHash
}

function Get-DevDecodeHash {
    $path = Get-DevLicensePath
    $obj = Read-LicenseFile $path
    if (-not $obj) { return '' }
    if (-not $obj.licenseType -or ([string]$obj.licenseType).ToLower() -ne 'developer') { return '' }
    if (-not $obj.decodeKeyHash) { return '' }
    return [string]$obj.decodeKeyHash
}

# Returns 'owner', 'developer', or 'user' based on license files present.
# owner   = dev license with both devKeyHash AND decodeKeyHash
# developer = dev license with devKeyHash only (no decodeKeyHash)
# user    = no dev license (user license only)
function Get-LicenseTier {
    $devHash = Get-DevLicenseHash
    if ([string]::IsNullOrWhiteSpace($devHash)) { return 'user' }
    $decodeHash = Get-DevDecodeHash
    if (-not [string]::IsNullOrWhiteSpace($decodeHash)) { return 'owner' }
    return 'developer'
}

function Install-DevLicenseFromFile([string]$sourcePath) {
    if ([string]::IsNullOrWhiteSpace($sourcePath)) {
        throw 'No developer license file was selected.'
    }
    if (-not (Test-Path $sourcePath)) {
        throw "Developer license file was not found: $sourcePath"
    }
    if ([System.IO.Path]::GetExtension($sourcePath).ToLower() -ne '.json') {
        throw 'Developer license file must be a .json file.'
    }

    $obj = Read-LicenseFile $sourcePath
    if (-not $obj) {
        throw "Selected developer license file could not be read: $sourcePath"
    }
    if (-not $obj.licenseType -or ([string]$obj.licenseType).ToLower() -ne 'developer') {
        throw "Selected file is not a developer license: $sourcePath"
    }

    $targetPath = Get-DevLicensePath
    $targetDir = Split-Path -Parent $targetPath
    if ($targetDir -and -not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    $sourceFull = [System.IO.Path]::GetFullPath($sourcePath)
    $targetFull = [System.IO.Path]::GetFullPath($targetPath)
    if ($sourceFull -ieq $targetFull) {
        return $targetPath
    }

    Copy-Item -Path $sourcePath -Destination $targetPath -Force
    return $targetPath
}

function Show-DevLicenseImportDialog {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $installedPath = ''
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Import Developer License'
    $dlg.Size = New-Object System.Drawing.Size(620, 340)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Import your developer license JSON by entering a full path, or use the Desktop helper, then click Import.'
    $lbl.Location = New-Object System.Drawing.Point(16, 16)
    $lbl.Size = New-Object System.Drawing.Size(580, 36)
    $dlg.Controls.Add($lbl)

    $dropPanel = New-Object System.Windows.Forms.Panel
    $dropPanel.Location = New-Object System.Drawing.Point(16, 58)
    $dropPanel.Size = New-Object System.Drawing.Size(580, 120)
    $dropPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $dropPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 249)
    $dropPanel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $dlg.Controls.Add($dropPanel)

    $dropLabel = New-Object System.Windows.Forms.Label
    $dropLabel.Text = "Enter full path below and click Import`n`nTip: use Desktop Dev Key"
    $dropLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $dropLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $dropLabel.Font = New-Object System.Drawing.Font('Segoe UI', 11)
    $dropLabel.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
    $dropLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $dropPanel.Controls.Add($dropLabel)

    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Location = New-Object System.Drawing.Point(16, 192)
    $txtPath.Size = New-Object System.Drawing.Size(580, 24)
    $txtPath.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $dlg.Controls.Add($txtPath)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = 'Drag-and-drop and Browse are disabled for compatibility. Use path import helpers.'
    $status.Location = New-Object System.Drawing.Point(16, 224)
    $status.Size = New-Object System.Drawing.Size(580, 22)
    $dlg.Controls.Add($status)

    $btnImport = New-Object System.Windows.Forms.Button
    $btnImport.Text = 'Import'
    $btnImport.Location = New-Object System.Drawing.Point(206, 258)
    $btnImport.Size = New-Object System.Drawing.Size(90, 30)
    $dlg.Controls.Add($btnImport)

    $btnDesktop = New-Object System.Windows.Forms.Button
    $btnDesktop.Text = 'Desktop Dev Key'
    $btnDesktop.Location = New-Object System.Drawing.Point(306, 258)
    $btnDesktop.Size = New-Object System.Drawing.Size(120, 30)
    $dlg.Controls.Add($btnDesktop)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(436, 258)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $dlg.Controls.Add($btnCancel)

    $importAction = {
        param([string]$candidatePath)

        try {
            $installed = Install-DevLicenseFromFile $candidatePath
            $script:devLicenseImportedThisRun = $true
            $status.Text = "Developer license installed to: $installed"
            $dlg.Tag = $installed
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dlg.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                $_.Exception.Message,
                'Developer License Import',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
    }.GetNewClosure()

    $desktopAction = {
        param($sender, $e)

        $desktop = [Environment]::GetFolderPath('Desktop')
        $candidate = Join-Path $desktop 'Dev Key.json'
        if (Test-Path $candidate) {
            $txtPath.Text = $candidate
            $status.Text = 'Desktop Dev Key path selected. Click Import to continue.'
        } else {
            $status.Text = 'Desktop Dev Key.json not found. Paste a path and click Import.'
        }
    }.GetNewClosure()

    $cancelAction = {
        param($sender, $e)
        $dlg.Tag = 'cancel'
        $dlg.Close()
    }.GetNewClosure()

    $importFromPathAction = {
        param($sender, $e)

        $candidatePath = [string]$txtPath.Text
        if ([string]::IsNullOrWhiteSpace($candidatePath)) {
            [System.Windows.Forms.MessageBox]::Show(
                'Paste a full path to a developer license JSON file, then click Import.',
                'Developer License Import',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        & $importAction $candidatePath
    }.GetNewClosure()

    $btnImport.Add_Click($importFromPathAction)
    $btnDesktop.Add_Click($desktopAction)
    $btnCancel.Add_Click($cancelAction)

    $null = $dlg.ShowDialog()
    $installedPath = [string]$dlg.Tag
    $dlg.Dispose()
    return (-not [string]::IsNullOrWhiteSpace($installedPath))
}

function Ensure-DeveloperLicensePresentInteractive {
    $script:devLicenseImportedThisRun = $false
    $path = Get-DevLicensePath
    $obj = Read-LicenseFile $path
    if ($obj -and $obj.licenseType -and ([string]$obj.licenseType).ToLower() -eq 'developer') {
        return $true
    }

    return (Show-DevLicenseImportDialog)
}

function Compute-StringHash([string]$s) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes     = [System.Text.Encoding]::UTF8.GetBytes($s)
        $hashBytes = $sha.ComputeHash($bytes)
        return ($hashBytes | ForEach-Object { $_.ToString('x2') }) -join ''
    } finally {
        $sha.Dispose()
    }
}

function Invoke-DevKeyPrompt {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $expectedHash = Get-DevLicenseHash

    if ([string]::IsNullOrWhiteSpace($expectedHash)) {
        $devPath = Get-DevLicensePath
        [System.Windows.Forms.MessageBox]::Show(
            "Developer license file is required for lock features and was not found/valid:`n$devPath",
            'Developer Access',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $false
    }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Developer Access - STEELMETTLE THC Systems Integrator'
    $dlg.Size = New-Object System.Drawing.Size(440, 180)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Enter developer key to unlock production lock controls:'
    $lbl.Location = New-Object System.Drawing.Point(16, 18)
    $lbl.Size = New-Object System.Drawing.Size(400, 22)
    $dlg.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(16, 48)
    $txt.Size = New-Object System.Drawing.Size(400, 26)
    $txt.PasswordChar = [char]0x2022
    $dlg.Controls.Add($txt)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Unlock'
    $btnOk.Location = New-Object System.Drawing.Point(228, 92)
    $btnOk.Size = New-Object System.Drawing.Size(90, 30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.AcceptButton = $btnOk
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(328, 92)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.CancelButton = $btnCancel
    $dlg.Controls.Add($btnCancel)

    $result  = $dlg.ShowDialog()
    $entered = $txt.Text
    $dlg.Dispose()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $false }
    $hash = Compute-StringHash $entered
    return ($hash.ToLower() -eq $expectedHash.ToLower())
}

function Invoke-DecodeKeyPrompt {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $expectedHash = Get-DevDecodeHash

    # Presence of decodeKeyHash in a valid developer license is treated as authorized.
    if (-not [string]::IsNullOrWhiteSpace($expectedHash)) {
        return $true
    }

    if ([string]::IsNullOrWhiteSpace($expectedHash)) {
        $devPath = Get-DevLicensePath
        [System.Windows.Forms.MessageBox]::Show(
            "Developer license file is required to export decoded Arduino and STM32 payloads and was not found/valid:`n$devPath`n`nEnsure dev.license.json contains a 'decodeKeyHash' field.",
            'Firmware Decode - Stage 2',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $false
    }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Firmware Decode Key - STEELMETTLE THC Systems Integrator'
    $dlg.Size = New-Object System.Drawing.Size(440, 180)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Enter decode key to export decoded Arduino and STM32 payloads:'
    $lbl.Location = New-Object System.Drawing.Point(16, 18)
    $lbl.Size = New-Object System.Drawing.Size(400, 22)
    $dlg.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(16, 48)
    $txt.Size = New-Object System.Drawing.Size(400, 26)
    $txt.PasswordChar = [char]0x2022
    $dlg.Controls.Add($txt)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Unlock'
    $btnOk.Location = New-Object System.Drawing.Point(228, 92)
    $btnOk.Size = New-Object System.Drawing.Size(90, 30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.AcceptButton = $btnOk
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(328, 92)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.CancelButton = $btnCancel
    $dlg.Controls.Add($btnCancel)

    $result  = $dlg.ShowDialog()
    $entered = $txt.Text
    $dlg.Dispose()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $false }
    $hash = Compute-StringHash $entered
    if ($hash.ToLower() -ne $expectedHash.ToLower()) {
        [System.Windows.Forms.MessageBox]::Show(
            'Incorrect decode key. Decoded payload export aborted.',
            'Firmware Decode - Stage 2',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $false
    }
    return $true
}

function Get-DecodedFirmwareBytes([string]$path) {
    $raw = Get-Content -Path $path -Raw -ErrorAction SilentlyContinue
    if ($raw -and $raw.TrimStart().StartsWith('SMTFW:')) {
        return Decode-TransferPayload -raw $raw -prefix 'SMTFW:' -scope 'firmware'
    }
    return [System.IO.File]::ReadAllBytes($path)
}

function Get-DecodedFirmwareText([string]$path) {
    $raw = Get-Content -Path $path -Raw -ErrorAction SilentlyContinue
    if ($raw -and $raw.TrimStart().StartsWith('SMTFW:')) {
        return [System.Text.Encoding]::UTF8.GetString((Decode-TransferPayload -raw $raw -prefix 'SMTFW:' -scope 'firmware'))
    }
    return $raw
}

function Export-DecodedFirmwarePayloads {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    if (-not $script:devModeUnlocked) {
        throw 'Developer mode must be unlocked before exporting decoded payloads.'
    }

    if (-not $script:decodeModeUnlocked) {
        if (-not (Invoke-DecodeKeyPrompt)) {
            throw 'Decode export aborted: invalid Stage 2 key.'
        }
        $script:decodeModeUnlocked = $true
    }

    $detected = Detect-Device
    if (-not $detected -or ($detected.Type -notin @('STM32','Arduino'))) {
        throw 'No device to export from. Connect STEELMETTLE Forge or STEELMETTLE THC Core and retry.'
    }

    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Choose a folder for decoded Arduino and STM32 payload export'
    $dlg.ShowNewFolderButton = $true
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $exportRoot = Join-Path $dlg.SelectedPath ("STEELMETTLE_Decoded_Firmware_" + $stamp)
    $arduinoSrc = Join-Path $baseDir 'Arduino'
    $stm32File = if ($config -and $config.stm32 -and $config.stm32.firmwareBinary) { Join-Path $baseDir $config.stm32.firmwareBinary } else { Join-Path $baseDir 'STM32\STM32MX_V2.bin' }

    New-Item -ItemType Directory -Path $exportRoot | Out-Null
    $arduinoOut = Join-Path $exportRoot 'Arduino'
    $stm32Out = Join-Path $exportRoot 'STM32'
    New-Item -ItemType Directory -Path $arduinoOut | Out-Null
    New-Item -ItemType Directory -Path $stm32Out | Out-Null

    if (Test-Path $arduinoSrc) {
        Get-ChildItem $arduinoSrc -File | Where-Object { $_.Extension -in '.ino','.cpp','.h','.c' } | ForEach-Object {
            $dest = Join-Path $arduinoOut $_.Name
            $raw = Get-Content -Path $_.FullName -Raw -ErrorAction SilentlyContinue
            if ($raw -and $raw.TrimStart().StartsWith('SMTFW:')) {
                Set-Content -Path $dest -Value (Get-DecodedFirmwareText $_.FullName) -Encoding UTF8
            } else {
                Copy-Item -Path $_.FullName -Destination $dest -Force
            }
        }
    }

    if (-not (Test-Path $stm32File)) {
        throw "STM32 firmware file not found: $stm32File"
    }

    [System.IO.File]::WriteAllBytes((Join-Path $stm32Out ([System.IO.Path]::GetFileName($stm32File))), (Get-DecodedFirmwareBytes $stm32File))
    return $exportRoot
}

function Normalize-VersionString([string]$v) {
    if ([string]::IsNullOrWhiteSpace($v)) { return '0.0.0' }
    $x = $v.Trim()
    if ($x.StartsWith('v', [System.StringComparison]::OrdinalIgnoreCase)) {
        $x = $x.Substring(1)
    }
    return $x
}

function Try-ParseVersion([string]$v) {
    $n = Normalize-VersionString $v
    $ver = $null
    if ([System.Version]::TryParse($n, [ref]$ver)) {
        return $ver
    }
    return $null
}

function Test-NewerVersion([string]$currentVersion, [string]$latestVersion) {
    $cur = Try-ParseVersion $currentVersion
    $lat = Try-ParseVersion $latestVersion
    if (-not $cur -or -not $lat) { return $false }
    return ($lat -gt $cur)
}

function Get-GitHubLatestRelease {
    param(
        [string]$Repo
    )

    if ([string]::IsNullOrWhiteSpace($Repo)) {
        throw 'Update repo is not configured (config.update.repo).'
    }

    $url = "https://api.github.com/repos/$Repo/releases/latest"
    $headers = @{ 'User-Agent' = 'STEELMETTLE-THC-Systems-Integrator' }
    try {
        return Invoke-RestMethod -Uri $url -Headers $headers -Method Get -TimeoutSec 20
    } catch {
        return $null
    }
}

function Download-File {
    param(
        [string]$Url,
        [string]$OutFile
    )
    $headers = @{ 'User-Agent' = 'STEELMETTLE-THC-Systems-Integrator' }
    Invoke-WebRequest -Uri $Url -OutFile $OutFile -Headers $headers -TimeoutSec 120
}

function Check-For-AppUpdate {
    $u = Get-UpdateConfig
    if (-not $u.enabled) {
        return @{
            status = 'disabled'
            message = 'App updates are disabled in configuration.'
            currentVersion = Normalize-VersionString $u.currentVersion
            latestVersion = Normalize-VersionString $u.currentVersion
            isUpdateAvailable = $false
            releaseName = ''
            assetName = ''
            assetUrl = ''
            releasePage = ''
        }
    }

    if ([string]::IsNullOrWhiteSpace([string]$u.repo)) {
        return @{
            status = 'misconfigured'
            message = 'Update repository is not configured.'
            currentVersion = Normalize-VersionString $u.currentVersion
            latestVersion = Normalize-VersionString $u.currentVersion
            isUpdateAvailable = $false
            releaseName = ''
            assetName = ''
            assetUrl = ''
            releasePage = ''
        }
    }

    $release = Get-GitHubLatestRelease -Repo $u.repo
    if (-not $release) {
        $cur = Normalize-VersionString $u.currentVersion
        return @{
            status = 'no-release'
            message = 'No published release metadata found for update channel.'
            currentVersion = $cur
            latestVersion = $cur
            isUpdateAvailable = $false
            releaseName = ''
            assetName = ''
            assetUrl = ''
            releasePage = ''
        }
    }
    $latestTag = [string]$release.tag_name
    $latestVersion = Normalize-VersionString $latestTag
    $currentVersion = Normalize-VersionString $u.currentVersion

    $asset = $null
    if ($release.assets -and $release.assets.Count -gt 0) {
        if (-not [string]::IsNullOrWhiteSpace($u.assetNamePattern)) {
            foreach ($a in $release.assets) {
                if ([string]$a.name -like "*$($u.assetNamePattern)*") {
                    $asset = $a
                    break
                }
            }
        }
        if (-not $asset) {
            foreach ($a in $release.assets) {
                if ([string]$a.name -like '*.exe') {
                    $asset = $a
                    break
                }
            }
        }
        if (-not $asset) {
            $asset = $release.assets[0]
        }
    }

    return @{
        status = 'ok'
        message = ''
        currentVersion = $currentVersion
        latestVersion = $latestVersion
        isUpdateAvailable = (Test-NewerVersion -currentVersion $currentVersion -latestVersion $latestVersion)
        releaseName = [string]$release.name
        assetName = if ($asset) { [string]$asset.name } else { '' }
        assetUrl = if ($asset) { [string]$asset.browser_download_url } else { '' }
        releasePage = [string]$release.html_url
    }
}

function Get-FirmwareUpdateConfig {
    $f = if ($config -and $config.firmwareUpdate) { $config.firmwareUpdate } else { $null }
    return @{
        checkOnConnect      = if ($f -and $null -ne $f.checkOnConnect) { [bool]$f.checkOnConnect } else { $false }
        checkOnLaunch       = if ($f -and $null -ne $f.checkOnLaunch) { [bool]$f.checkOnLaunch } else { $false }
        forgeEnabled        = if ($f -and $f.forge -and $null -ne $f.forge.enabled) { [bool]$f.forge.enabled } else { $false }
        forgeRepo           = if ($f -and $f.forge -and $f.forge.repo) { [string]$f.forge.repo } else { '' }
        forgeAssetPattern   = if ($f -and $f.forge -and $f.forge.assetPattern) { [string]$f.forge.assetPattern } else { '' }
        forgeCurrentVersion = if ($f -and $f.forge -and $f.forge.currentVersion) { [string]$f.forge.currentVersion } else { '0.0.0' }
        coreEnabled         = if ($f -and $f.core -and $null -ne $f.core.enabled) { [bool]$f.core.enabled } else { $false }
        coreRepo            = if ($f -and $f.core -and $f.core.repo) { [string]$f.core.repo } else { '' }
        coreAssetPattern    = if ($f -and $f.core -and $f.core.assetPattern) { [string]$f.core.assetPattern } else { '' }
        coreCurrentVersion  = if ($f -and $f.core -and $f.core.currentVersion) { [string]$f.core.currentVersion } else { '0.0.0' }
    }
}

function Check-For-FirmwareUpdate {
    param(
        [ValidateSet('Forge','Core')][string]$Target,
        [switch]$Force
    )
    $fw = Get-FirmwareUpdateConfig
    if (-not $Force -and -not $fw.checkOnConnect -and -not $fw.checkOnLaunch) { return $null }
    $enabled = if ($Target -eq 'Forge') { $fw.forgeEnabled }        else { $fw.coreEnabled }
    if (-not $enabled) { return $null }
    $repo    = if ($Target -eq 'Forge') { $fw.forgeRepo }           else { $fw.coreRepo }
    $pattern = if ($Target -eq 'Forge') { $fw.forgeAssetPattern }   else { $fw.coreAssetPattern }
    $curVer  = if ($Target -eq 'Forge') { $fw.forgeCurrentVersion } else { $fw.coreCurrentVersion }
    if ([string]::IsNullOrWhiteSpace($repo)) { return $null }
    $release   = Get-GitHubLatestRelease -Repo $repo
    $latestTag = [string]$release.tag_name
    $asset     = $null
    if ($release.assets -and $release.assets.Count -gt 0) {
        if (-not [string]::IsNullOrWhiteSpace($pattern)) {
            foreach ($a in $release.assets) {
                if ([string]$a.name -like "*$pattern*") { $asset = $a; break }
            }
        }
        if (-not $asset) { $asset = $release.assets[0] }
    }
    return @{
        target            = $Target
        currentVersion    = Normalize-VersionString $curVer
        latestVersion     = Normalize-VersionString $latestTag
        isUpdateAvailable = (Test-NewerVersion -currentVersion $curVer -latestVersion $latestTag)
        assetName         = if ($asset) { [string]$asset.name } else { '' }
        assetUrl          = if ($asset) { [string]$asset.browser_download_url } else { '' }
    }
}

function Download-And-RunAppUpdate {
    param(
        [string]$Url,
        [string]$FileName
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        throw 'No update download URL provided.'
    }

    $updatesDir = Join-Path $baseDir 'updates'
    if (-not (Test-Path $updatesDir)) {
        New-Item -ItemType Directory -Path $updatesDir | Out-Null
    }

    $safeName = if ([string]::IsNullOrWhiteSpace($FileName)) { 'update-installer.exe' } else { $FileName }
    $outFile = Join-Path $updatesDir $safeName
    Set-UiStage 'Downloading update package...' 70
    Download-File -Url $Url -OutFile $outFile

    Log "Update package downloaded: $outFile"
    $currentPid = $PID
    Start-Process -FilePath $outFile -ArgumentList @('--update', '--wait-pid', "$currentPid") -WorkingDirectory $updatesDir | Out-Null
    Start-Sleep -Milliseconds 800
}

function Invoke-ConfiguredCommand {
    param(
        [string]$CommandLine,
        [string]$LogFileName,
        [int]$TimeoutSeconds = 45
    )

    if ([string]::IsNullOrWhiteSpace($CommandLine)) {
        throw 'Command is not configured.'
    }

    $effectiveTimeout = if ($TimeoutSeconds -gt 0) { $TimeoutSeconds } else { 45 }
    $resolvedCommandLine = Resolve-ConfiguredCommandLine $CommandLine
    $logFile = Join-Path $logDir $LogFileName
    $tmpStdOut = Join-Path $env:TEMP ("smt-cmd-out-" + [guid]::NewGuid().ToString('N') + '.log')
    $tmpStdErr = Join-Path $env:TEMP ("smt-cmd-err-" + [guid]::NewGuid().ToString('N') + '.log')

    Log "Executing configured command: $resolvedCommandLine"
    Push-Location $baseDir
    try {
        $proc = Start-Process -FilePath 'cmd.exe' `
            -ArgumentList @('/c', $resolvedCommandLine) `
            -WorkingDirectory $baseDir `
            -RedirectStandardOutput $tmpStdOut `
            -RedirectStandardError $tmpStdErr `
            -NoNewWindow `
            -PassThru

        $deadline = (Get-Date).AddSeconds($effectiveTimeout)
        while (-not $proc.HasExited) {
            if ((Get-Date) -ge $deadline) {
                try { $proc.Kill() } catch {}
                $timeoutMessage = "Command timed out after $effectiveTimeout second(s). Check the device connection and COM port, then retry."
                Add-Content -Path $logFile -Value $timeoutMessage
                throw $timeoutMessage
            }
            try { [System.Windows.Forms.Application]::DoEvents() } catch {}
            Start-Sleep -Milliseconds 100
        }

        $proc.WaitForExit()
        $stdOut = if (Test-Path $tmpStdOut) { Get-Content -Path $tmpStdOut -Raw -ErrorAction SilentlyContinue } else { '' }
        $stdErr = if (Test-Path $tmpStdErr) { Get-Content -Path $tmpStdErr -Raw -ErrorAction SilentlyContinue } else { '' }
        $combined = (($stdOut, $stdErr | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine).Trim()

        if ([string]::IsNullOrWhiteSpace($combined)) {
            Set-Content -Path $logFile -Value '' -Encoding UTF8
        } else {
            Set-Content -Path $logFile -Value $combined -Encoding UTF8
        }

        if ($proc.ExitCode -ne 0) {
            $tail = if (Test-Path $logFile) { ((Get-Content -Path $logFile -Tail 8 -ErrorAction SilentlyContinue) -join [Environment]::NewLine).Trim() } else { '' }
            if ([string]::IsNullOrWhiteSpace($tail)) {
                throw "Command failed with exit code $($proc.ExitCode)"
            }
            throw "Command failed with exit code $($proc.ExitCode): $tail"
        }
    } finally {
        foreach ($tmpFile in @($tmpStdOut, $tmpStdErr)) {
            if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
        }
        Pop-Location
    }
}

function Apply-ProductionLock {
    param(
        [ValidateSet('Forge','Core')]
        [string]$Target
    )

    $sec = Get-SecurityConfig
    if (-not $sec.enabled) {
        throw 'Production lock is disabled in config.security.enabled.'
    }

    $cmd = if ($Target -eq 'Forge') { $sec.forgeLockCommand } else { $sec.coreLockCommand }
    if ([string]::IsNullOrWhiteSpace($cmd)) {
        throw "$Target lock command is not configured in config.security."
    }

    # Substitute {PORT} with detected device serial port (avrdude Core commands need the COM port)
    if ($cmd -match '\{PORT\}') {
        $detected = Detect-Device
        $port = if ($detected -and $detected.Port) { [string]$detected.Port } else { '' }
        if ([string]::IsNullOrWhiteSpace($port)) {
            throw 'Lock command requires {PORT} substitution but no device serial port was detected. Connect the device and retry.'
        }
        $cmd = $cmd.Replace('{PORT}', $port)
    }

    $timeoutSeconds = 45
    if ($sec.PSObject.Properties['commandTimeoutSeconds'] -and [int]$sec.commandTimeoutSeconds -gt 0) {
        $timeoutSeconds = [int]$sec.commandTimeoutSeconds
    }

    $lockLogName = "security-lock-" + $Target.ToLower() + ".log"
    try {
        Invoke-ConfiguredCommand -CommandLine $cmd -LogFileName $lockLogName -TimeoutSeconds $timeoutSeconds
    } catch {
        # avrdude often exits non-zero even on success; verify from log output
        $lockLogPath = Join-Path $logDir $lockLogName
        $lockLogText = if (Test-Path $lockLogPath) { Get-Content -Path $lockLogPath -Raw -ErrorAction SilentlyContinue } else { '' }
        $lockVerified = $lockLogText -match '(?im)1\s+verified' -and $lockLogText -match '(?im)Avrdude done\.'
        if (-not $lockVerified) {
            throw
        }
        Log "$Target production lock command exited non-zero but avrdude verified the write."
    }
    Clear-TargetPendingLock -Target $Target
    Log "$Target production lock command executed successfully."
}


function Set-ObjectProperty([object]$obj, [string]$name, $value) {
    if ($obj.PSObject.Properties[$name]) {
        $obj.$name = $value
    } else {
        $obj | Add-Member -NotePropertyName $name -NotePropertyValue $value
    }
}

function Save-Stm32AiOptions {
    param(
        [bool]$Enabled,
        [bool]$PredictEnable,
        [int]$PredictWindowSize,
        [double]$LearningRate,
        [double]$ThresholdV,
        [double]$ConfidenceMin
    )

    if (Test-Path $configPath) {
        $cfg = Get-Content -Path $configPath -Raw | ConvertFrom-Json
    } else {
        $cfg = [pscustomobject]@{}
    }

    if (-not $cfg.PSObject.Properties['stm32']) {
        $cfg | Add-Member -NotePropertyName 'stm32' -NotePropertyValue ([pscustomobject]@{})
    }
    if (-not $cfg.stm32.PSObject.Properties['ai']) {
        $cfg.stm32 | Add-Member -NotePropertyName 'ai' -NotePropertyValue ([pscustomobject]@{})
    }

    Set-ObjectProperty $cfg.stm32.ai 'enabled' $Enabled
    Set-ObjectProperty $cfg.stm32.ai 'predictEnable' $PredictEnable
    Set-ObjectProperty $cfg.stm32.ai 'predictWindowSize' $PredictWindowSize
    Set-ObjectProperty $cfg.stm32.ai 'learningRate' $LearningRate
    Set-ObjectProperty $cfg.stm32.ai 'thresholdV' $ThresholdV
    Set-ObjectProperty $cfg.stm32.ai 'confidenceMin' $ConfidenceMin

    $cfg | ConvertTo-Json -Depth 12 | Set-Content -Path $configPath -Encoding ascii
    $script:config = $cfg
    Log 'Forge AI-Tuning options saved to config.json'
}

function Set-DefineValueInText([string]$content, [string]$name, [string]$valueText) {
    $pattern = "(?m)^(\\s*#define\\s+" + [regex]::Escape($name) + "\\s+)([^\\r\\n]*?)(\\s*//.*)?$"
    if ($content -match $pattern) {
        return [regex]::Replace($content, $pattern, ('$1' + $valueText + '$3'))
    }
    return $content
}

function Apply-Stm32AiConfigOverrides {
    $ai = Get-Stm32AiOptions
    if (-not $ai.enabled) {
        Log 'Forge AI-Tuning options: disabled via config.stm32.ai.enabled = false'
        return
    }

    $sync = Get-SyncConfig
    $targetHeaders = @()
    if ($sync.stm32SourceConfigHeader) {
        $targetHeaders += $sync.stm32SourceConfigHeader
    }

    # Also patch local mirrored STM32 header if present in integrator payload.
    $localHeader = Join-Path $baseDir 'STM32\Core\Inc\THC_Config.h'
    if (Test-Path $localHeader) {
        $targetHeaders += $localHeader
    }

    $targetHeaders = $targetHeaders | Select-Object -Unique

    foreach ($headerPath in $targetHeaders) {
        if (-not (Test-Path $headerPath)) {
            Log "Forge AI-Tuning options: header not found, skipping ($headerPath)"
            continue
        }

        $raw = Get-Content -Path $headerPath -Raw
        $rawBefore = $raw
        $boolPredict = if ($ai.predictEnable) { 'true' } else { 'false' }
        $learningRateText = ([double]$ai.learningRate).ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $thresholdText = ([double]$ai.thresholdV).ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $confidenceText = ([double]$ai.confidenceMin).ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $raw = Set-DefineValueInText $raw 'THC_PREDICT_ENABLE' $boolPredict
        $raw = Set-DefineValueInText $raw 'THC_PREDICT_WINDOW_SIZE' ([string]$ai.predictWindowSize)
        $raw = Set-DefineValueInText $raw 'THC_PREDICT_LEARNING_RATE' ($learningRateText + 'f')
        $raw = Set-DefineValueInText $raw 'THC_PREDICT_THRESHOLD_V' ($thresholdText + 'f')
        $raw = Set-DefineValueInText $raw 'THC_PREDICT_CONFIDENCE_MIN' ($confidenceText + 'f')

        if ($raw -ne $rawBefore) {
            Set-Content -Path $headerPath -Value $raw -Encoding ascii
            Log "Forge AI-Tuning options applied to $headerPath"
        } else {
            Log "Forge AI-Tuning options: no matching predictive defines found in $headerPath"
        }
    }
}

function Sync-IntegratorPayload {
    Set-UiStage 'Syncing source folders...' 10
    $sync = Get-SyncConfig
    if (-not $sync.enabled) {
        Log 'Sync disabled via config.sync.enabled = false'
        Set-UiStage 'Sync disabled in config.' 10
        return
    }

    Log 'Sync step: checking source folders for updates'

    # Load previous sync hashes to detect actual changes
    $syncHashCacheFile = Join-Path $logDir 'sync_hashes.json'
    $previousHashes = Load-SyncHashes -CacheFile $syncHashCacheFile
    $currentHashes = @{}
    
    $syncedCount = 0

    # Sync Arduino firmware only if source has changed
    $arduinoDst = Join-Path $baseDir 'Arduino\THC_AUTOSET_Firmware'
    if (Test-Path $sync.arduinoSourceDir) {
        if (-not (Test-Path $arduinoDst)) { New-Item -ItemType Directory -Path $arduinoDst | Out-Null }

        # Get hash of the arduino source directory (using a marker file approach)
        # We'll check a key file to determine if sources changed
        $markerFile = Join-Path $sync.arduinoSourceDir 'THC_AUTOSET_Firmware.ino'
        if (Test-Path $markerFile) {
            $sourceHash = Get-FileSyncHash -Path $markerFile
            $hashKey = 'Arduino:main'
            
            if ($sourceHash -and $previousHashes[$hashKey] -eq $sourceHash) {
                Log "Sync skip: Arduino sources unchanged (hash: $($sourceHash.Substring(0, 8))...)"
            } elseif ($sourceHash) {
                # Source has changed, perform mirror sync
                $rc = & robocopy $sync.arduinoSourceDir $arduinoDst /MIR /R:1 /W:1 /XD .git .vscode 2>&1
                $rcText = ($rc | Out-String).Trim()
                if ($rcText) {
                    $rcText | Tee-Object -FilePath (Join-Path $logDir 'sync-arduino.log') | Out-Null
                }
                if ($LASTEXITCODE -gt 7) {
                    throw "Arduino sync failed with robocopy code $LASTEXITCODE"
                }
                Log "Sync step: Arduino payload updated from $($sync.arduinoSourceDir) (hash: $($sourceHash.Substring(0, 8))...)"
                Set-UiStage 'Arduino sources synced.' 20
                $syncedCount++
                $currentHashes[$hashKey] = $sourceHash
            }
        } else {
            # Fallback: no marker file found, do a sync anyway
            $rc = & robocopy $sync.arduinoSourceDir $arduinoDst /MIR /R:1 /W:1 /XD .git .vscode 2>&1
            if ($LASTEXITCODE -gt 7) {
                throw "Arduino sync failed with robocopy code $LASTEXITCODE"
            }
            Log "Sync step: Arduino payload updated from $($sync.arduinoSourceDir) (no marker hash)"
            Set-UiStage 'Arduino sources synced.' 20
            $syncedCount++
        }
    } else {
        Log "Sync step: Arduino source not found, skipping ($($sync.arduinoSourceDir))"
    }

    # Sync STM32 binary only if source has changed
    $stm32Dst = Join-Path $baseDir 'STM32\STM32MX_V2.bin'
    if (Test-Path $sync.stm32SourceBin) {
        $stm32DstDir = Split-Path $stm32Dst -Parent
        if (-not (Test-Path $stm32DstDir)) { New-Item -ItemType Directory -Path $stm32DstDir | Out-Null }
        
        $sourceHash = Get-FileSyncHash -Path $sync.stm32SourceBin
        $hashKey = 'STM32:binary'
        
        if ($sourceHash -and $previousHashes[$hashKey] -eq $sourceHash) {
            Log "Sync skip: STM32 binary unchanged (hash: $($sourceHash.Substring(0, 8))...)"
        } elseif ($sourceHash) {
            Copy-Item $sync.stm32SourceBin -Destination $stm32Dst -Force
            Log "Sync step: Forge firmware binary updated from $($sync.stm32SourceBin) (hash: $($sourceHash.Substring(0, 8))...)"
            Set-UiStage 'Forge firmware binary synced.' 25
            $syncedCount++
            $currentHashes[$hashKey] = $sourceHash
        }
    } else {
        Log "Sync step: Forge source binary not found, skipping ($($sync.stm32SourceBin))"
    }

    # Sync PoKeys macro only if source has changed
    $macroDst = Join-Path $baseDir 'PoKeys\POKEYS_POEXT_MACROS_CLEAN.txt'
    if (Test-Path $sync.pokeysMacroSourceFile) {
        $macroDstDir = Split-Path $macroDst -Parent
        if (-not (Test-Path $macroDstDir)) { New-Item -ItemType Directory -Path $macroDstDir | Out-Null }
        
        $sourceHash = Get-FileSyncHash -Path $sync.pokeysMacroSourceFile
        $hashKey = 'PoKeys:macro'
        
        if ($sourceHash -and $previousHashes[$hashKey] -eq $sourceHash) {
            Log "Sync skip: PoKeys macro unchanged (hash: $($sourceHash.Substring(0, 8))...)"
        } elseif ($sourceHash) {
            Copy-Item $sync.pokeysMacroSourceFile -Destination $macroDst -Force
            Log "Sync step: PoKeys macro updated from $($sync.pokeysMacroSourceFile) (hash: $($sourceHash.Substring(0, 8))...)"
            Set-UiStage 'PoKeys macro synced.' 30
            $syncedCount++
            $currentHashes[$hashKey] = $sourceHash
        }
    } else {
        Log "Sync step: PoKeys macro source not found, skipping ($($sync.pokeysMacroSourceFile))"
    }

    # Save updated hash cache for next run
    Save-SyncHashes -Hashes $currentHashes -CacheFile $syncHashCacheFile

    if ($syncedCount -eq 0) {
        Log "Sync step: No changes detected. All files up-to-date."
    } else {
        Log "Sync step: Updated $syncedCount component(s)."
    }

    # Keep STM32 predictive/AI option defines aligned with integrator config.
    try {
        Apply-Stm32AiConfigOverrides
        Set-UiStage 'Forge AI-Tuning options synced.' 28
    } catch {
        Log "WARNING: Failed applying Forge AI-Tuning options: $_"
    }
}

function Find-Executable($name) {
    $candidates = @(
        (Join-Path $baseDir $name),
        (Join-Path $toolsDir $name),
        (Join-Path (Join-Path $baseDir 'bin') $name)
    )
    foreach ($path in $candidates) {
        if (Test-Path $path) { return (Resolve-Path $path).Path }
    }
    return $null
}

function Find-CommandExecutable([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) {
        return $null
    }

    $candidateNames = [System.Collections.Generic.List[string]]::new()
    $candidateNames.Add($name)
    if (-not $name.EndsWith('.exe', [System.StringComparison]::OrdinalIgnoreCase)) {
        $candidateNames.Add($name + '.exe')
    }

    foreach ($candidate in $candidateNames) {
        $found = Find-Executable $candidate
        if ($found) {
            return $found
        }

        try {
            $cmd = Get-Command $candidate -CommandType Application -ErrorAction Stop | Select-Object -First 1
            if ($cmd -and $cmd.Source) { return $cmd.Source }
            if ($cmd -and $cmd.Path) { return $cmd.Path }
        } catch {}
    }

    if ($candidateNames -contains 'avrdude.exe') {
        $avrdudePatterns = @(
            (Join-Path $env:LOCALAPPDATA 'Arduino15\packages\arduino\tools\avrdude\*\bin\avrdude.exe'),
            (Join-Path $env:APPDATA 'Arduino15\packages\arduino\tools\avrdude\*\bin\avrdude.exe'),
            (Join-Path $env:ProgramFiles 'Arduino\hardware\tools\avr\bin\avrdude.exe'),
            (Join-Path ${env:ProgramFiles(x86)} 'Arduino\hardware\tools\avr\bin\avrdude.exe')
        )

        foreach ($pattern in $avrdudePatterns) {
            $match = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1
            if ($match) {
                return $match.FullName
            }
        }
    }

    return $null
}

function Resolve-ConfiguredCommandLine([string]$CommandLine) {
    if ([string]::IsNullOrWhiteSpace($CommandLine)) {
        return $CommandLine
    }

    $trimmed = $CommandLine.Trim()
    if ($trimmed -notmatch '^(?:"([^"]+)"|([^\s]+))(.*)$') {
        return $CommandLine
    }

    $exeToken = if ($matches[1]) { $matches[1] } else { $matches[2] }
    $remainder = $matches[3]

    if ([System.IO.Path]::IsPathRooted($exeToken) -or $exeToken -match '[\\/]') {
        return $CommandLine
    }

    $resolvedExe = Find-CommandExecutable $exeToken
    if ($resolvedExe) {
        return ('"{0}"{1}' -f $resolvedExe, $remainder)
    }

    return $CommandLine
}

function Get-SerialPorts() {
    Get-WmiObject Win32_SerialPort | Select-Object DeviceID, Description, PNPDeviceID
}

function Show-WelcomeScreen {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Legacy PowerShell hosts (v3) have repeated paint/runtime quirks; skip custom splash there.
    if ($PSVersionTable -and $PSVersionTable.PSVersion -and $PSVersionTable.PSVersion.Major -le 3) {
        return [System.Windows.Forms.DialogResult]::OK
    }

    $uiCfg = Get-UiConfig

    # Form - borderless custom chrome
    $splash = New-Object System.Windows.Forms.Form
    $splash.Text = ''
    $splash.Width  = 620
    $splash.Height = 460
    $splash.StartPosition   = 'CenterScreen'
    $splash.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $splash.MaximizeBox     = $false
    $splash.MinimizeBox     = $false
    $splash.BackColor       = [System.Drawing.Color]::FromArgb(14, 18, 22)

    if ($uiCfg.appIcon -and (Test-Path $uiCfg.appIcon)) {
        try { $splash.Icon = New-Object System.Drawing.Icon($uiCfg.appIcon) } catch {}
    }

    # Drag-to-move
    $script:_wDragStart = [System.Drawing.Point]::Empty
    $splash.Add_MouseDown({ param($s,$e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $script:_wDragStart = $e.Location
        }
    })
    $splash.Add_MouseMove({ param($s,$e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $splash.Left += $e.X - $script:_wDragStart.X
            $splash.Top  += $e.Y - $script:_wDragStart.Y
        }
    })

    # Top accent bar - amber gradient
    $accentTop = New-Object System.Windows.Forms.Panel
    $accentTop.Dock   = [System.Windows.Forms.DockStyle]::Top
    $accentTop.Height = 5
    $accentTop.BackColor = [System.Drawing.Color]::FromArgb(210, 160, 35)
    $splash.Controls.Add($accentTop)

    # Left accent strip - amber fade down
    $accentLeft = New-Object System.Windows.Forms.Panel
    $accentLeft.Width    = 5
    $accentLeft.Height   = 455
    $accentLeft.Location = [System.Drawing.Point]::new(0, 5)
    $accentLeft.BackColor = [System.Drawing.Color]::FromArgb(175, 125, 20)
    $splash.Controls.Add($accentLeft)

    # Close (x) button
    $btnX = New-Object System.Windows.Forms.Label
    $btnX.Text      = 'x'
    $btnX.Width     = 36
    $btnX.Height    = 30
    $btnX.Location  = [System.Drawing.Point]::new(578, 8)
    $btnX.Font      = New-Object System.Drawing.Font('Segoe UI', 15)
    $btnX.ForeColor = [System.Drawing.Color]::FromArgb(80, 100, 110)
    $btnX.BackColor = [System.Drawing.Color]::Transparent
    $btnX.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $btnX.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnX.Add_MouseEnter({ $btnX.ForeColor = [System.Drawing.Color]::FromArgb(230, 70, 50) })
    $btnX.Add_MouseLeave({ $btnX.ForeColor = [System.Drawing.Color]::FromArgb(80, 100, 110) })
    $btnX.Add_Click({ $splash.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $splash.Close() })
    $splash.Controls.Add($btnX)

    # Company name - large white
    $lblCompany = New-Object System.Windows.Forms.Label
    $lblCompany.AutoSize  = $false
    $lblCompany.Width     = 580
    $lblCompany.Height    = 64
    $lblCompany.Location  = [System.Drawing.Point]::new(20, 52)
    $lblCompany.Text      = 'STEELMETTLE LLC'
    $lblCompany.Font      = New-Object System.Drawing.Font('Segoe UI', 30, [System.Drawing.FontStyle]::Bold)
    $lblCompany.ForeColor = [System.Drawing.Color]::FromArgb(252, 250, 245)
    $lblCompany.BackColor = [System.Drawing.Color]::Transparent
    $lblCompany.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $splash.Controls.Add($lblCompany)

    # Product line - amber bold
    $lblProduct = New-Object System.Windows.Forms.Label
    $lblProduct.AutoSize  = $false
    $lblProduct.Width     = 580
    $lblProduct.Height    = 30
    $lblProduct.Location  = [System.Drawing.Point]::new(20, 116)
    $lblProduct.Text      = 'SYSTEMS  INTEGRATOR'
    $lblProduct.Font      = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $lblProduct.ForeColor = [System.Drawing.Color]::FromArgb(218, 168, 30)
    $lblProduct.BackColor = [System.Drawing.Color]::Transparent
    $lblProduct.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $splash.Controls.Add($lblProduct)

    # Rule with amber center glow
    $rule = New-Object System.Windows.Forms.Panel
    $rule.Width    = 500
    $rule.Height   = 2
    $rule.Location = [System.Drawing.Point]::new(60, 158)
    $rule.BackColor = [System.Drawing.Color]::FromArgb(185, 140, 28)
    $splash.Controls.Add($rule)

    # Welcome headline
    $lblWelcome = New-Object System.Windows.Forms.Label
    $lblWelcome.AutoSize  = $false
    $lblWelcome.Width     = 580
    $lblWelcome.Height    = 38
    $lblWelcome.Location  = [System.Drawing.Point]::new(20, 170)
    $lblWelcome.Text      = 'Welcome to STEELMETTLE LLC Systems Integrator'
    $lblWelcome.Font      = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $lblWelcome.ForeColor = [System.Drawing.Color]::FromArgb(228, 225, 218)
    $lblWelcome.BackColor = [System.Drawing.Color]::Transparent
    $lblWelcome.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $splash.Controls.Add($lblWelcome)

    # Sub-description
    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.AutoSize  = $false
    $lblDesc.Width     = 580
    $lblDesc.Height    = 24
    $lblDesc.Location  = [System.Drawing.Point]::new(20, 212)
    $lblDesc.Text      = 'CNC Torch Height Control | STM32 + Arduino | PoKeys Macro Integration'
    $lblDesc.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    $lblDesc.ForeColor = [System.Drawing.Color]::FromArgb(110, 135, 150)
    $lblDesc.BackColor = [System.Drawing.Color]::Transparent
    $lblDesc.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $splash.Controls.Add($lblDesc)

    # Feature pills
    $pillLabels = @('  Sync  ', '  Detect  ', '  Flash  ', '  Deploy  ')
    $pillX = 68
    foreach ($pt in $pillLabels) {
        $p = New-Object System.Windows.Forms.Label
        $p.Text      = $pt
        $p.AutoSize  = $false
        $p.Width     = 102
        $p.Height    = 28
        $p.Location  = [System.Drawing.Point]::new($pillX, 252)
        $p.Font      = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
        $p.ForeColor = [System.Drawing.Color]::FromArgb(205, 168, 35)
        $p.BackColor = [System.Drawing.Color]::FromArgb(28, 36, 22)
        $p.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $p.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $splash.Controls.Add($p)
        $pillX += 114
    }

    # Get started button
    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.Text     = '  GET STARTED  >'
    $btnStart.Width    = 210
    $btnStart.Height   = 48
    $btnStart.Location = [System.Drawing.Point]::new(205, 330)
    $btnStart.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnStart.FlatAppearance.BorderSize  = 1
    $btnStart.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 155, 20)
    $btnStart.BackColor  = [System.Drawing.Color]::FromArgb(195, 148, 18)
    $btnStart.ForeColor  = [System.Drawing.Color]::FromArgb(16, 18, 22)
    $btnStart.Font       = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $btnStart.Cursor     = [System.Windows.Forms.Cursors]::Hand
    $btnStart.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnStart.Add_MouseEnter({
        $btnStart.BackColor = [System.Drawing.Color]::FromArgb(240, 200, 55)
        $btnStart.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(255, 220, 90)
    })
    $btnStart.Add_MouseLeave({
        $btnStart.BackColor = [System.Drawing.Color]::FromArgb(195, 148, 18)
        $btnStart.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 155, 20)
    })
    $splash.AcceptButton = $btnStart
    $splash.Controls.Add($btnStart)

    # Footer tagline
    $lblFooter = New-Object System.Windows.Forms.Label
    $lblFooter.AutoSize  = $false
    $lblFooter.Width     = 580
    $lblFooter.Height    = 22
    $lblFooter.Location  = [System.Drawing.Point]::new(20, 426)
    $lblFooter.Text      = 'Precision Controls | Industrial Grade | Built for the Shop Floor'
    $lblFooter.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
    $lblFooter.ForeColor = [System.Drawing.Color]::FromArgb(52, 66, 74)
    $lblFooter.BackColor = [System.Drawing.Color]::Transparent
    $lblFooter.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $splash.Controls.Add($lblFooter)

    $dialogResult = $splash.ShowDialog()
    $splash.Dispose()
    return $dialogResult
}

function Show-GUI {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

    # Give this process a unique App User Model ID so Windows treats it as its
    # own distinct app in the taskbar (shows the form's custom icon, not PS icon)
    Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public class TaskbarAppId {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern void SetCurrentProcessExplicitAppUserModelID(string AppID);
}
'@
    [TaskbarAppId]::SetCurrentProcessExplicitAppUserModelID('SteelMettle.THCSystemsIntegrator')

    $welcomeResult = [System.Windows.Forms.DialogResult]::OK
    try {
        $welcomeResult = Show-WelcomeScreen
    } catch {
        Log "WARNING: Welcome screen failed to render: $_"
    }

    if ($welcomeResult -ne [System.Windows.Forms.DialogResult]::OK) {
        Log 'Welcome screen closed by user; exiting app before main window startup.'
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'STEELMETTLE THC Systems Integrator'
    $form.Width = 780
    $form.Height = 736
    $form.MinimumSize = New-Object System.Drawing.Size(780, 736)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.MaximizeBox = $true
    $form.ShowIcon = $true
    $form.ShowInTaskbar = $true

    $uiCfg = Get-UiConfig
    if ($uiCfg.appIcon -and (Test-Path $uiCfg.appIcon)) {
        try {
            $form.Icon = New-Object System.Drawing.Icon($uiCfg.appIcon)
        } catch {
            Log "WARNING: Failed to load app icon: $($uiCfg.appIcon)"
        }
    }

    $accentSteel = [System.Drawing.Color]::FromArgb(230, 44, 56, 62)
    $accentOlive = [System.Drawing.Color]::FromArgb(230, 88, 84, 40)
    $accentCopper = [System.Drawing.Color]::FromArgb(230, 122, 58, 34)
    $textPrimary = [System.Drawing.Color]::FromArgb(245, 244, 240)
    $textSecondary = [System.Drawing.Color]::FromArgb(220, 214, 204)
    if ($uiCfg.splashImage -and (Test-Path $uiCfg.splashImage)) {
        try {
            $form.BackgroundImage = [System.Drawing.Image]::FromFile($uiCfg.splashImage)
            $form.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
        } catch {
            Log "WARNING: Failed to load splash image: $($uiCfg.splashImage)"
        }
    }

    $overlay = New-Object System.Windows.Forms.Panel
    $overlay.Width = 740
    $overlay.Height = 630
    $overlay.Location = [System.Drawing.Point]::new(18, 18)
    $overlay.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $overlay.BackColor = [System.Drawing.Color]::FromArgb(175, 8, 10, 10)
    $form.Controls.Add($overlay)

    $title = New-Object System.Windows.Forms.Label
    $title.AutoSize = $false
    $title.Width = 700
    $title.Height = 44
    $title.Location = [System.Drawing.Point]::new(20, 18)
    $title.Text = 'STEELMETTLE THC Systems Integrator'
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 20, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $textPrimary
    $title.BackColor = [System.Drawing.Color]::Transparent
    $overlay.Controls.Add($title)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.AutoSize = $false
    $subtitle.Width = 700
    $subtitle.Height = 26
    $subtitle.Location = [System.Drawing.Point]::new(22, 62)
    $subtitle.Text = 'One-click sync, detect, flash, and deploy (Core + Forge)'
    $subtitle.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $subtitle.ForeColor = $textSecondary
    $subtitle.BackColor = [System.Drawing.Color]::Transparent
    $overlay.Controls.Add($subtitle)

    $status = New-Object System.Windows.Forms.Label
    $status.AutoSize = $false
    $status.Width = 700
    $status.Height = 46
    $status.Location = [System.Drawing.Point]::new(22, 112)
    $status.Text = ''
    $status.Font = New-Object System.Drawing.Font('Segoe UI', 11)
    $status.ForeColor = $textPrimary
    $status.BackColor = [System.Drawing.Color]::Transparent
    $overlay.Controls.Add($status)

    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Width = 700
    $progress.Height = 24
    $progress.Minimum = 0
    $progress.Maximum = 100
    $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progress.Value = 0
    $progress.Location = [System.Drawing.Point]::new(22, 162)
    $overlay.Controls.Add($progress)

    function Style-ActionButton([System.Windows.Forms.Button]$btn) {
        $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btn.FlatAppearance.BorderSize = 1
        $btn.FlatAppearance.BorderColor = $accentOlive
        $btn.BackColor = $accentSteel
        $btn.ForeColor = $textPrimary
        $btn.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
        $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    }

    function Set-UiProgress([string]$text, [int]$value) {
        $status.Text = $text
        if ($value -ge 0) {
            $progress.Value = [Math]::Max($progress.Minimum, [Math]::Min($progress.Maximum, $value))
        }
        $overlay.Refresh()
        $form.Refresh()
    }

    $script:UiProgressCallback = ${function:Set-UiProgress}

    $aiOptions = Get-Stm32AiOptions

    $panelHome = New-Object System.Windows.Forms.Panel
    $panelHome.Width = 700
    $panelHome.Height = 390
    $panelHome.Location = [System.Drawing.Point]::new(22, 196)
    $panelHome.BackColor = [System.Drawing.Color]::FromArgb(45, 14, 16, 16)
    $overlay.Controls.Add($panelHome)

    $panelStm32Step1 = New-Object System.Windows.Forms.Panel
    $panelStm32Step1.Width = 700
    $panelStm32Step1.Height = 344
    $panelStm32Step1.Location = [System.Drawing.Point]::new(22, 196)
    $panelStm32Step1.BackColor = [System.Drawing.Color]::FromArgb(45, 14, 16, 16)
    $panelStm32Step1.Visible = $false
    $overlay.Controls.Add($panelStm32Step1)

    $panelStm32Step2 = New-Object System.Windows.Forms.Panel
    $panelStm32Step2.Width = 700
    $panelStm32Step2.Height = 344
    $panelStm32Step2.Location = [System.Drawing.Point]::new(22, 196)
    $panelStm32Step2.BackColor = [System.Drawing.Color]::FromArgb(45, 14, 16, 16)
    $panelStm32Step2.Visible = $false
    $overlay.Controls.Add($panelStm32Step2)

    if ($uiCfg.forgeScreenImage -and (Test-Path $uiCfg.forgeScreenImage)) {
        try {
            $panelStm32Step2.BackgroundImage = [System.Drawing.Image]::FromFile($uiCfg.forgeScreenImage)
            $panelStm32Step2.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
        } catch {
            Log "WARNING: Failed to load Forge screen background: $($uiCfg.forgeScreenImage)"
        }
    }

    # Keep advanced AI tuning text readable even when the Forge background image is bright.
    $aiReadableBack = [System.Drawing.Color]::FromArgb(24, 28, 30)

    $chkAiEnabled = New-Object System.Windows.Forms.CheckBox
    $chkAiEnabled.Text = 'Enable STM32 AI integration options'
    $chkAiEnabled.Checked = [bool]$aiOptions.enabled
    $chkAiEnabled.ForeColor = $textPrimary
    $chkAiEnabled.BackColor = [System.Drawing.Color]::Transparent
    $chkAiEnabled.Location = [System.Drawing.Point]::new(20, 72)
    $chkAiEnabled.Width = 320
    $panelStm32Step1.Controls.Add($chkAiEnabled)

    $chkPredict = New-Object System.Windows.Forms.CheckBox
    $chkPredict.Text = 'Enable predictive voltage monitoring'
    $chkPredict.Checked = [bool]$aiOptions.predictEnable
    $chkPredict.ForeColor = $textPrimary
    $chkPredict.BackColor = [System.Drawing.Color]::Transparent
    $chkPredict.Location = [System.Drawing.Point]::new(20, 100)
    $chkPredict.Width = 320
    $panelStm32Step1.Controls.Add($chkPredict)

    $lblProfile = New-Object System.Windows.Forms.Label
    $lblProfile.Text = 'AI preset selection:'
    $lblProfile.ForeColor = $textPrimary
    $lblProfile.BackColor = [System.Drawing.Color]::Transparent
    $lblProfile.Location = [System.Drawing.Point]::new(20, 132)
    $lblProfile.Width = 130
    $panelStm32Step1.Controls.Add($lblProfile)

    $cmbProfile = New-Object System.Windows.Forms.ComboBox
    $cmbProfile.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbProfile.Items.Add('Conservative')
    [void]$cmbProfile.Items.Add('Balanced')
    [void]$cmbProfile.Items.Add('Aggressive')
    [void]$cmbProfile.Items.Add('Custom')
    $cmbProfile.SelectedItem = 'Balanced'
    $cmbProfile.Location = [System.Drawing.Point]::new(160, 128)
    $cmbProfile.Width = 190
    $panelStm32Step1.Controls.Add($cmbProfile)

    $lblWin = New-Object System.Windows.Forms.Label
    $lblWin.Text = 'Predict window size (2-100):'
    $lblWin.ForeColor = $textPrimary
    $lblWin.BackColor = $aiReadableBack
    $lblWin.Location = [System.Drawing.Point]::new(20, 62)
    $lblWin.Width = 220
    $panelStm32Step2.Controls.Add($lblWin)

    $txtWin = New-Object System.Windows.Forms.TextBox
    $txtWin.Location = [System.Drawing.Point]::new(248, 58)
    $txtWin.Width = 90
    $txtWin.Text = [string]$aiOptions.predictWindowSize
    $panelStm32Step2.Controls.Add($txtWin)

    $lblLearn = New-Object System.Windows.Forms.Label
    $lblLearn.Text = 'Learning rate (0.0-1.0):'
    $lblLearn.ForeColor = $textPrimary
    $lblLearn.BackColor = $aiReadableBack
    $lblLearn.Location = [System.Drawing.Point]::new(20, 94)
    $lblLearn.Width = 220
    $panelStm32Step2.Controls.Add($lblLearn)

    $txtLearn = New-Object System.Windows.Forms.TextBox
    $txtLearn.Location = [System.Drawing.Point]::new(248, 90)
    $txtLearn.Width = 90
    $txtLearn.Text = ([double]$aiOptions.learningRate).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $panelStm32Step2.Controls.Add($txtLearn)

    $lblThreshold = New-Object System.Windows.Forms.Label
    $lblThreshold.Text = 'Predict threshold V (0.1-50.0):'
    $lblThreshold.ForeColor = $textPrimary
    $lblThreshold.BackColor = $aiReadableBack
    $lblThreshold.Location = [System.Drawing.Point]::new(20, 126)
    $lblThreshold.Width = 220
    $panelStm32Step2.Controls.Add($lblThreshold)

    $txtThreshold = New-Object System.Windows.Forms.TextBox
    $txtThreshold.Location = [System.Drawing.Point]::new(248, 122)
    $txtThreshold.Width = 90
    $txtThreshold.Text = ([double]$aiOptions.thresholdV).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $panelStm32Step2.Controls.Add($txtThreshold)

    $lblConfidence = New-Object System.Windows.Forms.Label
    $lblConfidence.Text = 'Min confidence (0.0-1.0):'
    $lblConfidence.ForeColor = $textPrimary
    $lblConfidence.BackColor = $aiReadableBack
    $lblConfidence.Location = [System.Drawing.Point]::new(20, 158)
    $lblConfidence.Width = 220
    $panelStm32Step2.Controls.Add($lblConfidence)

    $txtConfidence = New-Object System.Windows.Forms.TextBox
    $txtConfidence.Location = [System.Drawing.Point]::new(248, 154)
    $txtConfidence.Width = 90
    $txtConfidence.Text = ([double]$aiOptions.confidenceMin).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $panelStm32Step2.Controls.Add($txtConfidence)

    $lblStep1Title = New-Object System.Windows.Forms.Label
    $lblStep1Title.Text = 'STEELMETTLE Forge: AI-Tuning Control - Step 1'
    $lblStep1Title.ForeColor = $textPrimary
    $lblStep1Title.BackColor = [System.Drawing.Color]::Transparent
    $lblStep1Title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
    $lblStep1Title.Location = [System.Drawing.Point]::new(20, 18)
    $lblStep1Title.Width = 520
    $panelStm32Step1.Controls.Add($lblStep1Title)

    $lblStep1Hint = New-Object System.Windows.Forms.Label
    $lblStep1Hint.Text = 'Choose default AI behavior, then flash directly or continue to advanced edits.'
    $lblStep1Hint.ForeColor = $textSecondary
    $lblStep1Hint.BackColor = [System.Drawing.Color]::Transparent
    $lblStep1Hint.Location = [System.Drawing.Point]::new(20, 44)
    $lblStep1Hint.Width = 520
    $panelStm32Step1.Controls.Add($lblStep1Hint)

    $lblStep2Title = New-Object System.Windows.Forms.Label
    $lblStep2Title.Text = 'STEELMETTLE Forge: AI-Tuning Control - Step 2'
    $lblStep2Title.ForeColor = $textPrimary
    $lblStep2Title.BackColor = $aiReadableBack
    $lblStep2Title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
    $lblStep2Title.Location = [System.Drawing.Point]::new(20, 18)
    $lblStep2Title.Width = 360
    $panelStm32Step2.Controls.Add($lblStep2Title)

    $lblStep2Hint = New-Object System.Windows.Forms.Label
    $lblStep2Hint.Text = 'Edit advanced AI values, save/apply, then flash rebuilt STM32 firmware.'
    $lblStep2Hint.ForeColor = $textSecondary
    $lblStep2Hint.BackColor = $aiReadableBack
    $lblStep2Hint.Location = [System.Drawing.Point]::new(20, 40)
    $lblStep2Hint.Width = 380
    $panelStm32Step2.Controls.Add($lblStep2Hint)

    function Set-AiProfile([string]$name) {
        switch ($name) {
            'Conservative' {
                $txtWin.Text = '12'
                $txtLearn.Text = '0.05'
                $txtThreshold.Text = '7.0'
                $txtConfidence.Text = '0.85'
            }
            'Balanced' {
                $txtWin.Text = '10'
                $txtLearn.Text = '0.1'
                $txtThreshold.Text = '5.0'
                $txtConfidence.Text = '0.7'
            }
            'Aggressive' {
                $txtWin.Text = '6'
                $txtLearn.Text = '0.2'
                $txtThreshold.Text = '3.5'
                $txtConfidence.Text = '0.55'
            }
        }
    }

    $cmbProfile.Add_SelectedIndexChanged({
        if ($cmbProfile.SelectedItem -and $cmbProfile.SelectedItem -ne 'Custom') {
            Set-AiProfile ([string]$cmbProfile.SelectedItem)
        }
    })

    function Get-ValidatedAiInputs {
        $win = 0
        if (-not [int]::TryParse($txtWin.Text, [ref]$win) -or $win -lt 2 -or $win -gt 100) {
            throw 'Predict window size must be an integer from 2 to 100.'
        }

        $learn = 0.0
        if (-not [double]::TryParse($txtLearn.Text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$learn) -or $learn -lt 0.0 -or $learn -gt 1.0) {
            throw 'Learning rate must be a number from 0.0 to 1.0 (use dot decimal).'
        }

        $thresh = 0.0
        if (-not [double]::TryParse($txtThreshold.Text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$thresh) -or $thresh -lt 0.1 -or $thresh -gt 50.0) {
            throw 'Predict threshold must be a number from 0.1 to 50.0 (V).'
        }

        $conf = 0.0
        if (-not [double]::TryParse($txtConfidence.Text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$conf) -or $conf -lt 0.0 -or $conf -gt 1.0) {
            throw 'Min confidence must be a number from 0.0 to 1.0.'
        }

        return @{
            enabled = [bool]$chkAiEnabled.Checked
            predictEnable = [bool]$chkPredict.Checked
            predictWindowSize = $win
            learningRate = $learn
            thresholdV = $thresh
            confidenceMin = $conf
        }
    }

    function Show-HomePanel {
        $panelHome.Visible = $true
        $panelStm32Step1.Visible = $false
        $panelStm32Step2.Visible = $false
        $subtitle.Text = 'One-click sync, detect, flash, and deploy (Core + Forge)'
    }

    function Show-Stm32Step1Panel {
        $panelHome.Visible = $false
        $panelStm32Step1.Visible = $true
        $panelStm32Step2.Visible = $false
        $subtitle.Text = 'STEELMETTLE Forge: AI-Tuning Control'
    }

    function Show-Stm32Step2Panel {
        $panelHome.Visible = $false
        $panelStm32Step1.Visible = $false
        $panelStm32Step2.Visible = $true
        $subtitle.Text = 'STEELMETTLE Forge: AI-Tuning Control'
    }

    function Enter-Stm32Flow {
        $detected = Detect-Device
        if ($detected -and $detected.Type -eq 'STM32') {
            Set-UiProgress "Forge board detected (STM32): $($detected.Desc)" 35
            Show-Stm32Step1Panel
            return
        }

        $choice = [System.Windows.Forms.MessageBox]::Show(
            'Forge board (STM32) not detected. Continue with Forge AI-Tuning setup anyway?',
            'Forge Selection',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
            Set-UiProgress 'Forge mode selected manually. Configure AI-Tuning below.' 30
            Show-Stm32Step1Panel
        }
    }

    function Prompt-And-RunUpdateFlow {
        $update = $null
        try {
            $update = Check-For-AppUpdate
        } catch {
            Set-UiProgress 'Unable to check for updates right now.' 100
            [System.Windows.Forms.MessageBox]::Show(
                "Could not check for updates right now.`n`n$_",
                'Update Check',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        if (-not $update) {
            Set-UiProgress 'Unable to check for updates right now.' 100
            return
        }

        if ($update.status -eq 'disabled') {
            Set-UiProgress 'App updates are disabled in this build.' 100
            [System.Windows.Forms.MessageBox]::Show(
                'App updates are disabled in this build.',
                'Update Check',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        if ($update.status -eq 'misconfigured') {
            Set-UiProgress 'Update configuration is missing.' 100
            [System.Windows.Forms.MessageBox]::Show(
                'Update repository is not configured. Please contact support.',
                'Update Check',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        if ($update.status -eq 'no-release') {
            Set-UiProgress 'App is up to date.' 100
            [System.Windows.Forms.MessageBox]::Show(
                "Your app is up to date.`nCurrent version: $($update.currentVersion)",
                'Update Check',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        if (-not $update.isUpdateAvailable) {
            Set-UiProgress 'App is up to date.' 100
            [System.Windows.Forms.MessageBox]::Show(
                "Your app is up to date.`nCurrent version: $($update.currentVersion)`nLatest version: $($update.latestVersion)",
                'Update Check',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $msg = "Update available: $($update.latestVersion)`nCurrent: $($update.currentVersion)`n`nThe app will close automatically so the installer can replace the running files.`n`nDownload and install now?"
        $choice = [System.Windows.Forms.MessageBox]::Show(
            $msg,
            'Update Available',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }

        Set-UiProgress 'Downloading integrator update package...' 65
        Download-And-RunAppUpdate -Url $update.assetUrl -FileName $update.assetName
        Set-UiProgress 'Update downloaded. Installer launched. Closing...' 100
        $script:isAppUpdateShutdown = $true
        $form.Close()
        [System.Windows.Forms.Application]::Exit()
        [System.Environment]::Exit(0)
    }

    function Prompt-And-RunFirmwareUpdateFlow([string]$target, [switch]$Force) {
        try {
            $fwParams = @{ Target = $target }
            if ($Force) { $fwParams['Force'] = $true }
            $fw = Check-For-FirmwareUpdate @fwParams
            if (-not $fw -or -not $fw.isUpdateAvailable) { return }
            $msg = "$target firmware update available: v$($fw.latestVersion)`nCurrent: v$($fw.currentVersion)`n`nDownload and stage updated $target firmware now?"
            $choice = [System.Windows.Forms.MessageBox]::Show(
                $msg, "$target Firmware Update",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            $subDir  = if ($target -eq 'Forge') { 'STM32' } else { 'Arduino' }
            $fwDir   = Join-Path $baseDir $subDir
            if (-not (Test-Path $fwDir)) { New-Item -ItemType Directory -Path $fwDir | Out-Null }
            $outFile = Join-Path $fwDir $fw.assetName
            Set-UiProgress "Downloading $target firmware update..." 55
            Download-File -Url $fw.assetUrl -OutFile $outFile
            Set-UiProgress "$target firmware staged for flash." 80
            Log "$target firmware update staged: $outFile"
        } catch {
            Log "WARNING: $target firmware update check failed: $_"
        }
    }

    function Confirm-And-ApplyProductionLock([string]$target) {
        $sec = Get-SecurityConfig
        if ($sec.requireConfirmation) {
            $warn = "Apply production lock for $target now?`n`nThis can block firmware readback and may require erase/reprovision to service devices."
            $ans = [System.Windows.Forms.MessageBox]::Show(
                $warn,
                "$target Production Lock",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }
        }

        Set-UiProgress "Applying $target production lock..." 80
        try {
            Apply-ProductionLock -Target $target
        } catch {
            throw
        }
        Set-UiProgress "$target production lock applied." 100
    }

    $btnAuto = New-Object System.Windows.Forms.Button
    $btnAuto.Text = 'Auto-Detect & Run'
    $btnAuto.Width = 342
    $btnAuto.Height = 46
    $btnAuto.Location = [System.Drawing.Point]::new(22, 16)
    Style-ActionButton $btnAuto
    $btnAuto.Add_Click({
        try {
            Run-Auto
            Set-UiProgress 'Done.' 100
        } catch {
            Set-UiProgress "Error: $_" 100
        }
    })
    $panelHome.Controls.Add($btnAuto)

    $btnSTM = New-Object System.Windows.Forms.Button
    $btnSTM.Text = 'Open STEELMETTLE Forge: AI-Tuning Control'
    $btnSTM.Width = 342
    $btnSTM.Height = 46
    $btnSTM.Location = [System.Drawing.Point]::new(380, 16)
    Style-ActionButton $btnSTM
    $btnSTM.Add_Click({
        try {
            Enter-Stm32Flow
        } catch {
            Set-UiProgress "Error: $_" 100
        }
    })
    $panelHome.Controls.Add($btnSTM)

    $btnArduino = New-Object System.Windows.Forms.Button
    $btnArduino.Text = 'Upload STEELMETTLE THC Core'
    $btnArduino.Width = 342
    $btnArduino.Height = 46
    $btnArduino.Location = [System.Drawing.Point]::new(22, 76)
    Style-ActionButton $btnArduino
    $btnArduino.Add_Click({
        try {
            Set-UiProgress 'Uploading STEELMETTLE THC Core...' 30
            Sync-IntegratorPayload
            Set-UiProgress 'Uploading STEELMETTLE THC Core...' 60
            Build-And-Flash-Arduino
            Set-UiProgress 'STEELMETTLE THC Core upload done.' 100
        } catch {
            Set-UiProgress "Error: $_" 100
        }
    })
    $panelHome.Controls.Add($btnArduino)

    $btnPokeys = New-Object System.Windows.Forms.Button
    $btnPokeys.Text = 'Deploy PoKeys'
    $btnPokeys.Width = 342
    $btnPokeys.Height = 46
    $btnPokeys.Location = [System.Drawing.Point]::new(380, 76)
    Style-ActionButton $btnPokeys
    $btnPokeys.Add_Click({
        try {
            Sync-IntegratorPayload
            Set-UiProgress 'Deploying PoKeys...' 60
            Deploy-PoKeys
            Set-UiProgress 'PoKeys done.' 100
        } catch {
            Set-UiProgress "Error: $_" 100
        }
    })
    $panelHome.Controls.Add($btnPokeys)

    $btnMachMacro = New-Object System.Windows.Forms.Button
    $btnMachMacro.Text = 'Install Mach3/Mach4 Feedrate Macros'
    $btnMachMacro.Width = 700
    $btnMachMacro.Height = 46
    $btnMachMacro.Location = [System.Drawing.Point]::new(0, 136)
    Style-ActionButton $btnMachMacro
    $btnMachMacro.BackColor = $accentOlive
    $btnMachMacro.Add_Click({
        try {
            if (-not (Ensure-MachMacroModuleLoaded)) {
                [System.Windows.Forms.MessageBox]::Show(
                    'Mach macro helper could not be loaded on this system. App startup is unaffected, but macro install is unavailable.',
                    'Mach Macro Helper Unavailable',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }

            $mach3Installed = if (Test-Path 'C:\Mach3') { $true } else { $false }
            $mach4Installed = if (Test-Path 'C:\Mach4') { $true } else { $false }

            if (-not $mach3Installed -and -not $mach4Installed) {
                [System.Windows.Forms.MessageBox]::Show(
                    'No Mach3 or Mach4 installation detected. Please install Mach3 or Mach4 first.',
                    'Mach Not Found',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }

            $options = @()
            if ($mach3Installed) { $options += 'Mach3' }
            if ($mach4Installed) { $options += 'Mach4' }

            if ($options.Count -eq 1) {
                $selectedVersion = $options[0]
            } else {
                $dlg = New-Object System.Windows.Forms.Form
                $dlg.Text = 'Select Mach Version'
                $dlg.Width = 320
                $dlg.Height = 160
                $dlg.StartPosition = 'CenterParent'
                $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
                $dlg.MaximizeBox = $false
                $dlg.MinimizeBox = $false

                $lbl = New-Object System.Windows.Forms.Label
                $lbl.Text = 'Choose which Mach version to install feedrate macro for:'
                $lbl.Location = [System.Drawing.Point]::new(16, 16)
                $lbl.Width = 280
                $lbl.Height = 30
                $dlg.Controls.Add($lbl)

                $cmbMach = New-Object System.Windows.Forms.ComboBox
                $cmbMach.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
                foreach ($opt in $options) { [void]$cmbMach.Items.Add($opt) }
                $cmbMach.SelectedIndex = 0
                $cmbMach.Location = [System.Drawing.Point]::new(16, 50)
                $cmbMach.Width = 280
                $dlg.Controls.Add($cmbMach)

                $btnOk = New-Object System.Windows.Forms.Button
                $btnOk.Text = 'Install'
                $btnOk.Location = [System.Drawing.Point]::new(176, 92)
                $btnOk.Width = 60
                $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $dlg.AcceptButton = $btnOk
                $dlg.Controls.Add($btnOk)

                $btnCancel = New-Object System.Windows.Forms.Button
                $btnCancel.Text = 'Cancel'
                $btnCancel.Location = [System.Drawing.Point]::new(244, 92)
                $btnCancel.Width = 60
                $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $dlg.CancelButton = $btnCancel
                $dlg.Controls.Add($btnCancel)

                if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
                    $dlg.Dispose()
                    return
                }
                $selectedVersion = $cmbMach.SelectedItem
                $dlg.Dispose()
            }

            Set-UiProgress "Installing $selectedVersion feedrate macro..." 50
            $result = Install-MachFeedrateMacro -MachVersion $selectedVersion
            if ($result.Success) {
                $statusMsg = if ($result.AutoConfigured) {
                    "$selectedVersion macro fully installed and configured at: $($result.MacroPath)"
                } else {
                    "$selectedVersion macro installed at: $($result.MacroPath) (manual config needed)"
                }
                Set-UiProgress $statusMsg 100
                [System.Windows.Forms.MessageBox]::Show(
                    "Macro installed to: $($result.MacrosFolder)`n`n$($result.Message)",
                    "$selectedVersion Macro Installation Complete",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
            } else {
                Set-UiProgress "Error installing $selectedVersion macro: $($result.Message)" 100
                [System.Windows.Forms.MessageBox]::Show(
                    "Installation failed: $($result.Message)",
                    "$selectedVersion Macro Installation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
            }
        } catch {
            Set-UiProgress "Error: $_" 100
        }
    })
    $panelHome.Controls.Add($btnMachMacro)

    $btnUpdate = New-Object System.Windows.Forms.Button
    $btnUpdate.Text = 'Check for Updates (App & Firmware)'
    $btnUpdate.Width = 700
    $btnUpdate.Height = 46
    $btnUpdate.Location = [System.Drawing.Point]::new(0, 182)
    Style-ActionButton $btnUpdate
    $btnUpdate.BackColor = $accentOlive
    $btnUpdate.Add_Click({
        try {
            Set-UiProgress 'Checking for updates...' 30
            Prompt-And-RunUpdateFlow
            Prompt-And-RunFirmwareUpdateFlow 'Forge' -Force
            Prompt-And-RunFirmwareUpdateFlow 'Core' -Force
        } catch {
            Set-UiProgress 'Unable to check updates right now.' 100
        }
    })
    $panelHome.Controls.Add($btnUpdate)

    $btnLockCore = New-Object System.Windows.Forms.Button
    $btnLockCore.Text = 'Apply STEELMETTLE THC Core Production Lock'
    $btnLockCore.Width = 700
    $btnLockCore.Height = 46
    $btnLockCore.Location = [System.Drawing.Point]::new(0, 228)
    Style-ActionButton $btnLockCore
    $btnLockCore.BackColor = $accentCopper
    $btnLockCore.Add_Click({
        try {
            Confirm-And-ApplyProductionLock -target 'Core'
        } catch {
            Set-UiProgress "Core lock error: $_" 100
        }
    })
    $btnLockCore.Visible = $false
    $panelHome.Controls.Add($btnLockCore)

    $btnDecodePayloads = New-Object System.Windows.Forms.Button
    $btnDecodePayloads.Text = 'Export Decoded Arduino + STM32 Payloads'
    $btnDecodePayloads.Width = 700
    $btnDecodePayloads.Height = 34
    $btnDecodePayloads.Location = [System.Drawing.Point]::new(0, 338)
    Style-ActionButton $btnDecodePayloads
    $btnDecodePayloads.BackColor = $accentSteel
    $btnDecodePayloads.Add_Click({
        try {
            $exportRoot = Export-DecodedFirmwarePayloads
            if ($exportRoot) {
                Set-UiProgress "Decoded payloads exported to $exportRoot" 100
            } else {
                Set-UiProgress 'Decoded payload export cancelled.' 100
            }
        } catch {
            Set-UiProgress "Decoded payload export error: $_" 100
        }
    })
    $btnDecodePayloads.Visible = $false
    $panelHome.Controls.Add($btnDecodePayloads)

    $btnDevAccess = New-Object System.Windows.Forms.Button
    $btnDevAccess.Text = '- Developer Access -'
    $btnDevAccess.Width = 700
    $btnDevAccess.Height = 26
    $btnDevAccess.Location = [System.Drawing.Point]::new(0, 352)
    $btnDevAccess.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDevAccess.FlatAppearance.BorderSize = 0
    $btnDevAccess.BackColor = $accentCopper
    $btnDevAccess.ForeColor = $textSecondary
    $btnDevAccess.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $btnDevAccess.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnDevAccess.Add_Click({
        try {
            if ($script:devModeUnlocked) {
                Set-UiProgress 'Developer mode is already active.' 100
                return
            }
            $hadDevLicenseBefore = -not [string]::IsNullOrWhiteSpace((Get-DevLicenseHash))
            if (-not (Ensure-DeveloperLicensePresentInteractive)) {
                Set-UiProgress 'Developer license import cancelled.' 100
                return
            }
            $hasDevLicenseNow = -not [string]::IsNullOrWhiteSpace((Get-DevLicenseHash))
            if (($script:devLicenseImportedThisRun -or (-not $hadDevLicenseBefore -and $hasDevLicenseNow)) -and $hasDevLicenseNow) {
                Set-UiProgress 'Developer license added. Restarting app...' 100
                Start-Sleep -Milliseconds 300
                $null = Restart-IntegratorApp
                $form.Close()
                [System.Windows.Forms.Application]::Exit()
                [System.Environment]::Exit(0)
                return
            }
            if (Invoke-DevKeyPrompt) {
                $script:devModeUnlocked = $true
                $tier = Get-LicenseTier
                $btnLockCore.Visible = $true
                $btnLockForge.Visible = $true
                $btnDecodePayloads.Visible = ($tier -eq 'owner')
                $btnDevAccess.Visible = $false
                Set-UiProgress 'Developer mode unlocked. Production lock controls are now visible.' 100
            } else {
                Set-UiProgress 'Developer access denied.' 100
            }
        } catch {
            Set-UiProgress "Developer access error: $_" 100
        }
    })
    $panelHome.Controls.Add($btnDevAccess)

    # If a valid developer license is present, unlock developer mode automatically.
    $licenseTier = Get-LicenseTier
    if ($licenseTier -eq 'owner') {
        $script:devModeUnlocked = $true
        $btnLockCore.Visible = $true
        $btnDecodePayloads.Visible = $true
        $btnDevAccess.Visible = $false
    } elseif ($licenseTier -eq 'developer') {
        $script:devModeUnlocked = $true
        $btnLockCore.Visible = $true
        $btnDecodePayloads.Visible = $false
        $btnDevAccess.Visible = $false
    } else {
        # user tier: no dev features, hide dev access button entirely
        $btnDevAccess.Visible = $false
    }

    $btnAi = New-Object System.Windows.Forms.Button
    $btnAi.Text = 'Upload Default STEELMETTLE Forge AI-Tuning + Flash'
    $btnAi.Width = 700
    $btnAi.Height = 46
    $btnAi.Location = [System.Drawing.Point]::new(0, 208)
    Style-ActionButton $btnAi
    $btnAi.BackColor = $accentOlive
    $btnAi.Add_Click({
        try {
            if ($cmbProfile.SelectedItem -and $cmbProfile.SelectedItem -ne 'Custom') {
                Set-AiProfile ([string]$cmbProfile.SelectedItem)
            }
            Set-UiProgress 'Applying default Forge AI-Tuning profile...' 40
            $vals = Get-ValidatedAiInputs
            Save-Stm32AiOptions -Enabled $vals.enabled -PredictEnable $vals.predictEnable -PredictWindowSize $vals.predictWindowSize -LearningRate $vals.learningRate -ThresholdV $vals.thresholdV -ConfidenceMin $vals.confidenceMin
            Apply-Stm32AiConfigOverrides
            Sync-IntegratorPayload
            Set-UiProgress 'Rebuilding and flashing Forge with selected defaults...' 70
            Build-And-Flash-STM32 -ForceBuild
            Set-UiProgress 'Forge flashed with default AI profile.' 100
        } catch {
            Set-UiProgress "Error: $_" 100
        }
    })
    $panelStm32Step1.Controls.Add($btnAi)

    $btnNextAdvanced = New-Object System.Windows.Forms.Button
    $btnNextAdvanced.Text = 'Next: STEELMETTLE Forge Advanced AI-Tuning'
    $btnNextAdvanced.Width = 340
    $btnNextAdvanced.Height = 42
    $btnNextAdvanced.Location = [System.Drawing.Point]::new(360, 128)
    Style-ActionButton $btnNextAdvanced
    $btnNextAdvanced.Add_Click({ Show-Stm32Step2Panel })
    $panelStm32Step1.Controls.Add($btnNextAdvanced)

    $btnBackFromStep1 = New-Object System.Windows.Forms.Button
    $btnBackFromStep1.Text = 'Back'
    $btnBackFromStep1.Width = 140
    $btnBackFromStep1.Height = 34
    $btnBackFromStep1.Location = [System.Drawing.Point]::new(560, 12)
    Style-ActionButton $btnBackFromStep1
    $btnBackFromStep1.BackColor = $accentCopper
    $btnBackFromStep1.Add_Click({ Show-HomePanel })
    $panelStm32Step1.Controls.Add($btnBackFromStep1)
    $btnBackFromStep1.BringToFront()

    $btnSaveApplyAdvanced = New-Object System.Windows.Forms.Button
    $btnSaveApplyAdvanced.Text = 'Save + Apply STEELMETTLE Forge Advanced AI-Tuning'
    $btnSaveApplyAdvanced.Width = 700
    $btnSaveApplyAdvanced.Height = 42
    $btnSaveApplyAdvanced.Location = [System.Drawing.Point]::new(0, 206)
    Style-ActionButton $btnSaveApplyAdvanced
    $btnSaveApplyAdvanced.BackColor = $accentOlive
    $btnSaveApplyAdvanced.Add_Click({
        try {
            Set-UiProgress 'Saving advanced Forge AI-Tuning profile...' 45
            $vals = Get-ValidatedAiInputs
            Save-Stm32AiOptions -Enabled $vals.enabled -PredictEnable $vals.predictEnable -PredictWindowSize $vals.predictWindowSize -LearningRate $vals.learningRate -ThresholdV $vals.thresholdV -ConfidenceMin $vals.confidenceMin
            Apply-Stm32AiConfigOverrides
            Set-UiProgress 'Advanced Forge AI-Tuning saved and applied.' 60
        } catch {
            Set-UiProgress "Error: $_" 100
        }
    })
    $panelStm32Step2.Controls.Add($btnSaveApplyAdvanced)

    $btnFlashAdvanced = New-Object System.Windows.Forms.Button
    $btnFlashAdvanced.Text = 'Flash STEELMETTLE Forge with Updated AI-Tuning'
    $btnFlashAdvanced.Width = 700
    $btnFlashAdvanced.Height = 42
    $btnFlashAdvanced.Location = [System.Drawing.Point]::new(0, 254)
    Style-ActionButton $btnFlashAdvanced
    $btnFlashAdvanced.BackColor = $accentCopper
    $btnFlashAdvanced.Add_Click({
        try {
            Set-UiProgress 'Preparing Forge flash with updated AI-Tuning...' 65
            $vals = Get-ValidatedAiInputs
            Save-Stm32AiOptions -Enabled $vals.enabled -PredictEnable $vals.predictEnable -PredictWindowSize $vals.predictWindowSize -LearningRate $vals.learningRate -ThresholdV $vals.thresholdV -ConfidenceMin $vals.confidenceMin
            Apply-Stm32AiConfigOverrides
            Sync-IntegratorPayload
            Build-And-Flash-STM32 -ForceBuild
            Set-UiProgress 'Forge flashed with updated AI-Tuning.' 100
        } catch {
            Set-UiProgress "Error: $_" 100
        }
    })
    $panelStm32Step2.Controls.Add($btnFlashAdvanced)

    $btnLockForge = New-Object System.Windows.Forms.Button
    $btnLockForge.Text = 'Apply STEELMETTLE Forge Production Lock'
    $btnLockForge.Width = 700
    $btnLockForge.Height = 34
    $btnLockForge.Location = [System.Drawing.Point]::new(0, 300)
    Style-ActionButton $btnLockForge
    $btnLockForge.BackColor = $accentCopper
    $btnLockForge.Add_Click({
        try {
            Confirm-And-ApplyProductionLock -target 'Forge'
        } catch {
            Set-UiProgress "Forge lock error: $_" 100
        }
    })
    $btnLockForge.Visible = $false
    if ($script:devModeUnlocked) {
        $btnLockForge.Visible = $true
    }
    $panelStm32Step2.Controls.Add($btnLockForge)

    $btnBackFromStep2 = New-Object System.Windows.Forms.Button
    $btnBackFromStep2.Text = 'Back to Step 1'
    $btnBackFromStep2.Width = 140
    $btnBackFromStep2.Height = 34
    $btnBackFromStep2.Location = [System.Drawing.Point]::new(560, 12)
    Style-ActionButton $btnBackFromStep2
    $btnBackFromStep2.BackColor = $accentCopper
    $btnBackFromStep2.Add_Click({ Show-Stm32Step1Panel })
    $panelStm32Step2.Controls.Add($btnBackFromStep2)

    $btnBackHomeFromStep2 = New-Object System.Windows.Forms.Button
    $btnBackHomeFromStep2.Text = 'Back to Home'
    $btnBackHomeFromStep2.Width = 140
    $btnBackHomeFromStep2.Height = 34
    $btnBackHomeFromStep2.Location = [System.Drawing.Point]::new(410, 12)
    Style-ActionButton $btnBackHomeFromStep2
    $btnBackHomeFromStep2.BackColor = $accentCopper
    $btnBackHomeFromStep2.Add_Click({ Show-HomePanel })
    $panelStm32Step2.Controls.Add($btnBackHomeFromStep2)
    $btnBackHomeFromStep2.BringToFront()
    $btnBackFromStep2.BringToFront()

    function Update-ResponsiveLayout {
        $contentLeft = 22
        $contentWidth = [Math]::Max(520, $overlay.ClientSize.Width - 44)

        $title.SetBounds(20, 18, [Math]::Max(400, $overlay.ClientSize.Width - 40), 44)
        $subtitle.SetBounds($contentLeft, 62, $contentWidth, 26)
        $status.SetBounds($contentLeft, 112, $contentWidth, 46)
        $progress.SetBounds($contentLeft, 162, $contentWidth, 24)

        if ($btnClose) {
            $btnCloseTop = [Math]::Max(200, $overlay.ClientSize.Height - 48)
            $btnClose.SetBounds($contentLeft, $btnCloseTop, $contentWidth, 40)
        }

        $panelTop = 196
        $panelBottom = if ($btnClose) { $btnClose.Top - 12 } else { $overlay.ClientSize.Height - 52 }
        $panelHeight = [Math]::Max(220, $panelBottom - $panelTop)

        $panelHome.SetBounds($contentLeft, $panelTop, $contentWidth, $panelHeight)
        $panelStm32Step1.SetBounds($contentLeft, $panelTop, $contentWidth, $panelHeight)
        $panelStm32Step2.SetBounds($contentLeft, $panelTop, $contentWidth, $panelHeight)

        $homeScaleY = [double]$panelHome.ClientSize.Height / 390.0
        $colGap = 16
        $colWidth = [int](($panelHome.ClientSize.Width - $colGap) / 2)
        $homePrimaryH = [int][Math]::Max(34, [Math]::Round(46 * $homeScaleY))
        $homeCompactH = [int][Math]::Max(24, [Math]::Round(34 * $homeScaleY))
        $homeDevH = [int][Math]::Max(20, [Math]::Round(26 * $homeScaleY))

        $btnAuto.SetBounds(0, [int][Math]::Round(16 * $homeScaleY), $colWidth, $homePrimaryH)
        $btnSTM.SetBounds($colWidth + $colGap, [int][Math]::Round(16 * $homeScaleY), $colWidth, $homePrimaryH)
        $btnArduino.SetBounds(0, [int][Math]::Round(76 * $homeScaleY), $colWidth, $homePrimaryH)
        $btnPokeys.SetBounds($colWidth + $colGap, [int][Math]::Round(76 * $homeScaleY), $colWidth, $homePrimaryH)

        $fullHomeWidth = $panelHome.ClientSize.Width
        $btnMachMacro.SetBounds(0, [int][Math]::Round(136 * $homeScaleY), $fullHomeWidth, $homePrimaryH)
        $btnUpdate.SetBounds(0, [int][Math]::Round(182 * $homeScaleY), $fullHomeWidth, $homePrimaryH)
        $btnLockCore.SetBounds(0, [int][Math]::Round(228 * $homeScaleY), $fullHomeWidth, $homePrimaryH)
        $btnDecodePayloads.SetBounds(0, [int][Math]::Round(338 * $homeScaleY), $fullHomeWidth, $homeCompactH)
        $btnDevAccess.SetBounds(0, [int][Math]::Round(352 * $homeScaleY), $fullHomeWidth, $homeDevH)

        $step1ScaleY = [double]$panelStm32Step1.ClientSize.Height / 344.0
        $fullStep1Width = $panelStm32Step1.ClientSize.Width
        $step1Btn42H = [int][Math]::Max(32, [Math]::Round(42 * $step1ScaleY))
        $step1Btn46H = [int][Math]::Max(34, [Math]::Round(46 * $step1ScaleY))
        $btnNextAdvanced.SetBounds(0, [int][Math]::Round(128 * $step1ScaleY), $fullStep1Width, $step1Btn42H)
        $btnAi.SetBounds(0, [int][Math]::Round(208 * $step1ScaleY), $fullStep1Width, $step1Btn46H)
        $btnBackFromStep1.Location = [System.Drawing.Point]::new([Math]::Max(0, $fullStep1Width - 140), [int][Math]::Round(12 * $step1ScaleY))

        $step2ScaleY = [double]$panelStm32Step2.ClientSize.Height / 344.0
        $fullStep2Width = $panelStm32Step2.ClientSize.Width
        $step2Btn42H = [int][Math]::Max(32, [Math]::Round(42 * $step2ScaleY))
        $step2Btn34H = [int][Math]::Max(24, [Math]::Round(34 * $step2ScaleY))
        $btnSaveApplyAdvanced.SetBounds(0, [int][Math]::Round(206 * $step2ScaleY), $fullStep2Width, $step2Btn42H)
        $btnFlashAdvanced.SetBounds(0, [int][Math]::Round(254 * $step2ScaleY), $fullStep2Width, $step2Btn42H)
        $btnLockForge.SetBounds(0, [int][Math]::Round(300 * $step2ScaleY), $fullStep2Width, $step2Btn34H)
        $backBtnTop = [int][Math]::Round(12 * $step2ScaleY)
        $btnBackFromStep2.Location = [System.Drawing.Point]::new([Math]::Max(0, $fullStep2Width - 140), $backBtnTop)
        $btnBackHomeFromStep2.Location = [System.Drawing.Point]::new([Math]::Max(0, $fullStep2Width - 290), $backBtnTop)
    }

    $form.Add_Resize({ Update-ResponsiveLayout })
    $overlay.Add_Resize({ Update-ResponsiveLayout })
    Update-ResponsiveLayout

    Show-HomePanel

    # One-time firmware update check on launch (non-continuous)
    $fwCfg = Get-FirmwareUpdateConfig
    if ($fwCfg.checkOnLaunch) {
        $form.Add_Shown({
            try {
                Prompt-And-RunFirmwareUpdateFlow 'Forge' -Force
                Prompt-And-RunFirmwareUpdateFlow 'Core' -Force
            } catch {
                Log "WARNING: Launch firmware update check failed: $_"
            }
        })
    }

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Width = 700
    $btnClose.Height = 40
    $btnClose.Location = [System.Drawing.Point]::new(22, 582)
    $btnClose.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClose.FlatAppearance.BorderSize = 1
    $btnClose.FlatAppearance.BorderColor = $accentOlive
    $btnClose.BackColor = $accentCopper
    $btnClose.ForeColor = $textPrimary
    $btnClose.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $form.Close() })
    $overlay.Controls.Add($btnClose)

    $updateCfg = Get-UpdateConfig
    $lblVersion = New-Object System.Windows.Forms.Label
    $lblVersion.Text = "v$($updateCfg.currentVersion)"
    $lblVersion.AutoSize = $true
    $lblVersion.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lblVersion.ForeColor = $textSecondary
    $lblVersion.BackColor = [System.Drawing.Color]::Transparent
    $lblVersion.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $lblVersion.Location = [System.Drawing.Point]::new(($overlay.ClientSize.Width - 70), ($overlay.ClientSize.Height - 20))
    $overlay.Controls.Add($lblVersion)
    $lblVersion.BringToFront()

    Update-ResponsiveLayout

    $form.Add_FormClosing({
        param($sender, $e)

        if ($script:isAppUpdateShutdown) {
            return
        }

        if (-not (Test-PendingDevLocksRequireAction)) {
            return
        }

        if (-not (Ensure-PendingDevLocksBeforeExit -OwnerForm $sender)) {
            $e.Cancel = $true
        }
    })

    [void][System.Windows.Forms.Application]::Run($form)
}

function Run-Auto {
    Sync-IntegratorPayload
    Set-UiStage 'Looking for device...' 35
    $device = Detect-Device
    if (-not $device) {
        throw 'No supported USB device detected.'
    }

    Set-UiStage "Found $($device.Type): $($device.Desc)" 45

    switch ($device.Type) {
        'STM32' { Build-And-Flash-STM32 }
        'Arduino' { Build-And-Flash-Arduino }
        'PoKeys' { Deploy-PoKeys }
        default { throw "Unsupported device type: $($device.Type)" }
    }
}

function Detect-Device() {
    $ports = Get-SerialPorts

    # Heuristic: ST-Link Virtual COM shows up as "STMicroelectronics" or "STLink".
    # Arduino shows up as "Arduino" or "CDC".
    # PoKeys uses "PoKeys" or "PoLabs" in the description.

    foreach ($port in $ports) {
        $desc = $port.Description
        if ($desc -match 'STMicroelectronics|STLink|ST-LINK') { return @{ Type='STM32'; Port=$port.DeviceID; Desc=$desc }}
        if ($desc -match 'Arduino|CH340|CP210|USB Serial') { return @{ Type='Arduino'; Port=$port.DeviceID; Desc=$desc }}
        if ($desc -match 'PoKeys|PoLabs') { return @{ Type='PoKeys'; Port=$port.DeviceID; Desc=$desc }}
    }

    return $null
}

function Build-And-Flash-STM32 {
    param(
        [switch]$ForceBuild
    )

    Set-UiStage 'Preparing STEELMETTLE Forge flash...' 60
    Log 'Starting STEELMETTLE Forge build + flash flow'
    $stm32Exe = Find-Executable 'STM32_Programmer_CLI.exe'
    if (-not $stm32Exe) {
        Log 'ERROR: STM32_Programmer_CLI.exe not found. Place it in this folder or tools/. '
        throw 'STM32_Programmer_CLI.exe missing.'
    }

    $stm32Dir = Join-Path $baseDir 'STM32'
    $binFile = if ($config -and $config.stm32 -and $config.stm32.firmwareBinary) { Join-Path $baseDir $config.stm32.firmwareBinary } else { Join-Path $stm32Dir 'STM32MX_V2.bin' }

    # Build when forced (AI options changed) or when binary is missing.
    if ($ForceBuild -or -not (Test-Path $binFile)) {
        $buildDir = if ($config -and $config.stm32 -and $config.stm32.buildDir) { Join-Path $baseDir $config.stm32.buildDir } else { Join-Path $stm32Dir 'Debug' }
        if (Test-Path $buildDir) {
            if ($ForceBuild) {
                Log "Force-building STEELMETTLE Forge firmware in $buildDir"
            } else {
                Log "Forge binary not found; attempting build in $buildDir"
            }
            Push-Location $buildDir
            & make 2>&1 | Tee-Object -FilePath (Join-Path $logDir 'build-stm32.log')
            Pop-Location
        }
    }

    if (-not (Test-Path $binFile)) {
        Log 'ERROR: Failed to locate Forge firmware binary. Place the .bin in the STM32 folder or provide build sources.'
        throw 'Missing firmware binary'
    }

    # Decode firmware binary if it was encoded for distribution (SMTFW: prefix)
    $tempBin = $null
    $flashBin = $binFile
    try {
        $binContent = Get-Content -Path $binFile -Raw -ErrorAction SilentlyContinue
        if ($binContent -and $binContent.TrimStart().StartsWith('SMTFW:')) {
                Log 'Encoded firmware detected - decoding payload for flash.'
            $binBytes = Get-DecodedFirmwareBytes $binFile
            $tempBin = [System.IO.Path]::GetTempFileName() + '.bin'
            [System.IO.File]::WriteAllBytes($tempBin, $binBytes)
            $flashBin = $tempBin
        }
        Log "Flashing STEELMETTLE Forge with binary: $binFile"
        Set-UiStage 'Found Forge board, installing firmware...' 80
        & $stm32Exe -c port=SWD -d $flashBin 2>&1 | Tee-Object -FilePath (Join-Path $logDir 'flash-stm32.log')
        if ($script:devModeUnlocked) {
            Mark-TargetPendingLock -Target 'Forge'
        }
        Log 'STEELMETTLE Forge flash complete.'
        Set-UiStage 'Forge firmware install complete.' 100
    } finally {
        if ($tempBin) { Remove-Item $tempBin -Force -ErrorAction SilentlyContinue }
    }
}

function Build-And-Flash-Arduino {
    param(
        [string]$ComPort
    )

    Set-UiStage 'Preparing STEELMETTLE THC Core build...' 60
    Log 'Starting STEELMETTLE THC Core build + flash flow'
    $arduino = Find-Executable 'arduino-cli.exe'
    if (-not $arduino) {
        Log 'ERROR: arduino-cli.exe not found. Place it in this folder or tools/.'
        throw 'arduino-cli.exe missing.'
    }

    $arduinoDir = Join-Path $baseDir 'Arduino'
    $sketchDir = if ($config -and $config.arduino -and $config.arduino.sketchDir) { Join-Path $baseDir $config.arduino.sketchDir } else { $arduinoDir }

    if (-not (Test-Path $sketchDir)) {
        Log "ERROR: Arduino sketch directory not found: $sketchDir"
        throw 'Arduino sketch missing.'
    }

    $fqbn = if ($config -and $config.arduino -and $config.arduino.fqbn) { $config.arduino.fqbn } else { 'arduino:avr:uno' }

    # Decode sketch sources if they were encoded for distribution (SMTFW: prefix)
    $tempSketchRoot = $null
    $tempSketchDir = $null
    $buildSketchDir = $sketchDir
    $sampleIno = Get-ChildItem $sketchDir -Filter '*.ino' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($sampleIno) {
        $inoContent = Get-Content $sampleIno.FullName -Raw -ErrorAction SilentlyContinue
        if ($inoContent -and $inoContent.TrimStart().StartsWith('SMTFW:')) {
            Log 'Encoded Arduino sketch detected - decoding payload for flash.'
            $tempSketchRoot = Join-Path $env:TEMP "smt_sketch_$(Get-Random)"
            $sketchFolderName = [System.IO.Path]::GetFileNameWithoutExtension($sampleIno.Name)
            $tempSketchDir = Join-Path $tempSketchRoot $sketchFolderName
            New-Item -ItemType Directory -Path $tempSketchDir -Force | Out-Null
            Get-ChildItem $sketchDir -Recurse -File | ForEach-Object {
                $fc = Get-Content $_.FullName -Raw -ErrorAction SilentlyContinue
                $relativePath = $_.FullName.Substring($sketchDir.Length).TrimStart([char]'\', [char]'/')
                $dp = Join-Path $tempSketchDir $relativePath
                $dpDir = Split-Path $dp -Parent
                if (-not (Test-Path $dpDir)) {
                    New-Item -ItemType Directory -Path $dpDir -Force | Out-Null
                }
                if ($fc -and $fc.TrimStart().StartsWith('SMTFW:')) {
                    $decodedText = Get-DecodedFirmwareText $_.FullName
                    [System.IO.File]::WriteAllText($dp, $decodedText, (New-Object System.Text.UTF8Encoding($false)))
                } else {
                    Copy-Item $_.FullName $dp -Force
                }
            }
            $buildSketchDir = $tempSketchDir
        }
    }

    try {
        Log "Compiling STEELMETTLE THC Core sketch in $sketchDir (board $fqbn)"
        Set-UiStage 'Compiling STEELMETTLE THC Core firmware...' 70
        $prevEap = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            & $arduino compile --fqbn $fqbn $buildSketchDir 2>&1 | Tee-Object -FilePath (Join-Path $logDir 'build-arduino.log')
            $compileExit = $LASTEXITCODE
        } finally {
            $ErrorActionPreference = $prevEap
        }
        if ($compileExit -ne 0) {
            throw "Arduino compile failed with exit code $compileExit"
        }

        if (-not $ComPort) {
            $detected = Detect-Device
            if ($detected -and $detected.Type -eq 'Arduino') {
                $ComPort = $detected.Port
            }
        }

        if (-not $ComPort) {
            $ComPort = Read-Host 'Enter Arduino COM port (e.g., COM3)'
        }

        Log "Uploading to $ComPort"
        Set-UiStage "Found Core board on $ComPort, installing firmware..." 85
        $prevEap = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            & $arduino upload -p $ComPort --fqbn $fqbn $buildSketchDir 2>&1 | Tee-Object -FilePath (Join-Path $logDir 'flash-arduino.log')
            $uploadExit = $LASTEXITCODE
        } finally {
            $ErrorActionPreference = $prevEap
        }
        if ($uploadExit -ne 0) {
            throw "Arduino upload failed with exit code $uploadExit"
        }
        if ($script:devModeUnlocked) {
            Mark-TargetPendingLock -Target 'Core'
        }
        Log 'STEELMETTLE THC Core upload complete.'
        Set-UiStage 'STEELMETTLE THC Core firmware install complete.' 100
    } finally {
        if ($tempSketchRoot) {
            Remove-Item $tempSketchRoot -Recurse -Force -ErrorAction SilentlyContinue
        } elseif ($tempSketchDir) {
            Remove-Item $tempSketchDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Deploy-PoKeys {
    Set-UiStage 'Preparing PoKeys deployment...' 60
    Log 'Starting PoKeys deployment flow'

    # Locate the CLI shim (pokeys-cli.ps1)
    $pokeysCliShim = $null
    $shimCandidates = @(
        (Join-Path $baseDir 'pokeys-cli.ps1'),
        (Join-Path $toolsDir 'pokeys-cli.ps1')
    )
    foreach ($candidate in $shimCandidates) {
        if (Test-Path $candidate) {
            $pokeysCliShim = (Resolve-Path $candidate).Path
            break
        }
    }

    if ($pokeysCliShim) {
        Log "Deploying PoKeys via DLL: $pokeysCliShim --deploy"
        Set-UiStage 'Configuring PoKeys device...' 70
        $deployLog = Join-Path $logDir 'pokeys-deploy.log'
        $output = & powershell -NoProfile -ExecutionPolicy Bypass -File $pokeysCliShim --deploy 2>&1
        $output | Out-File -FilePath $deployLog -Encoding utf8
        $output | ForEach-Object { Log "  $_" }
        if ($LASTEXITCODE -eq 0) {
            Log 'PoKeys deployment complete - device configured and bridge started.'
            Set-UiStage 'PoKeys deployment complete.' 100
        } else {
            Log "WARNING: PoKeys deploy exited with code $LASTEXITCODE"
            Set-UiStage 'PoKeys deployment failed - check log.' 100
        }
        return
    }

    # Fallback: no CLI shim found
    Log 'WARNING: pokeys-cli.ps1 not found. PoKeys deployment requires the tools folder.'
    Set-UiStage 'PoKeys CLI not found.' 100
}

# --- main ---
Log '=== STEELMETTLE THC Systems Integrator started ==='

try {
    Ensure-UserLicensePresent
} catch {
    Log "ERROR: $_"
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            [string]$_,
            'License Required',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } catch {
        # If WinForms is not available, console output is enough.
    }
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}

if ($Gui) {
    try {
        Show-GUI
        Log '=== STEELMETTLE THC Systems Integrator finished (GUI) ==='
        exit 0
    } catch {
        Log "ERROR: GUI startup failed: $_"
        try {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show(
                "GUI startup failed:`n`n$_`n`nSee logs at: $logDir",
                'STEELMETTLE THC Systems Integrator - Startup Error',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } catch {
            Write-Host "GUI startup failed: $_" -ForegroundColor Red
        }
        exit 1
    }
}

if ($Auto) {
    try {
        Run-Auto
        Log '=== STEELMETTLE THC Systems Integrator finished (Auto) ==='
        Write-Host 'Done.' -ForegroundColor Green
        exit 0
    } catch {
        Log "ERROR: $_"
        Write-Host "Error: $_" -ForegroundColor Red
        exit 1
    }
}

# Default (non-auto) behavior: detect and run automatically, prompting only when needed.
Sync-IntegratorPayload
$device = Detect-Device
if (-not $device) {
    Log 'No supported device detected. Please connect STEELMETTLE Forge, STEELMETTLE THC Core, or PoKeys and rerun.'
    Write-Host 'No supported device detected. Connect a board and rerun.' -ForegroundColor Yellow
    exit 1
}

Log "Detected device: $($device.Type) on $($device.Port) ($($device.Desc))"
Write-Host "Detected device: $($device.Type) ($($device.Desc))" -ForegroundColor Cyan

switch ($device.Type) {
    'STM32' { Build-And-Flash-STM32 }
    'Arduino' { Build-And-Flash-Arduino }
    'PoKeys' { Deploy-PoKeys }
    default {
        Write-Host "Detected unknown device type: $($device.Type)" -ForegroundColor Red
        Log "Unknown device type: $($device.Type)"
        exit 1
    }
}

Log '=== STEELMETTLE THC Systems Integrator finished ==='
Write-Host 'Done.' -ForegroundColor Green
