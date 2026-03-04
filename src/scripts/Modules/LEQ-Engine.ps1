# LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
# Licensed under GPL-3.0. See LICENSE file for details.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'


# =============================================================================
# Registry Constants (previously in Registry-Constants.ps1)
# =============================================================================

# Registry Paths
$script:REG_MMDEVICES_RENDER = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"

# Device Property GUIDs
# PKEY_Device_FriendlyName - the user-editable name shown in Sound settings (e.g., "Art Tune (Game)")
$script:PROP_DEVICE_FRIENDLY_NAME = "{a45c254e-df1c-4efd-8020-67d146a850e0},2"
# PKEY_Device_DeviceDesc - the driver/hardware description (e.g., "Sound Blaster GC7")
$script:PROP_DEVICE_DESC = "{b3f8fa53-0004-438e-9003-51a46e139bfc},6"
$script:PROP_DEVICE_FORMAT = "{f19f064d-082c-4e27-bc73-6882a1bb8e4c},0"
$script:PROP_ENHANCEMENT_TAB = "{1e94c58f-3e40-4ddb-b10c-a86d8b870a31},2"

# FX Property Base and Slots
$script:FX_PROPERTY_BASE = "{d04e05a6-594b-4fb6-a80d-01af5eed7d1d}"
$script:FX_SLOT_LFX = ",1"
$script:FX_SLOT_GFX = ",2"
$script:FX_SLOT_ENHANCEMENT = ",3"
$script:FX_SLOT_SFX = ",5"
$script:FX_SLOT_MFX = ",6"
# CompositeFX slots (same base GUID, higher slot numbers - Win10 1803+)
$script:FX_SLOT_COMPOSITE_SFX = ",13"
$script:FX_SLOT_COMPOSITE_MFX = ",14"
$script:FX_SLOT_COMPOSITE_EFX = ",15"

# APO GUIDs
$script:LEQ_APO_GUID = "{62dc1a93-ae24-464c-a43e-452f824c4250}"
$script:OTHER_APO_GUID = "{637c490d-eee3-4c0a-973f-371958802da2}"
$script:WM_GFX_APO_GUID = "{13AB3EBD-137E-4903-9D89-60BE8277FD17}"
$script:WM_LFX_APO_GUID = "{C9453E73-8C5C-4463-9984-AF8BAB2F5447}"
$script:ENHANCEMENT_TAB_GUID = "{5860E1C5-F95C-4a7a-8EC8-8AEF24F379A1}"
$script:EAPO_PRE_MIX_GUID = "{EACD2258-FCAC-4FF4-B36D-419E924A6D79}"
$script:EAPO_POST_MIX_GUID = "{EC1CC9CE-FAED-4822-828A-82A81A6F018F}"

# Enhancement DLL path (present on all Windows 10/11 systems)
$script:ENHANCEMENT_DLL = "C:\WINDOWS\System32\WMALFXGFXDSP.dll"

# E-APO Detection
$script:EAPO_CHILD_APO_PATH = "HKLM:\SOFTWARE\EqualizerAPO\Child APOs"

# Processing Modes property base - PKEY_SFX/MFX_ProcessingModes_Supported_For_Streaming
# These declare which audio processing modes the device's APOs support (e.g. Default mode).
# Must NOT be deleted - without them, the audio engine won't route audio through the APO chain.
$script:PROCESSING_MODES_PROPERTY_KEY = "{d3993a3f-99c2-4402-b5ec-a92a0367664b}"
$script:AUDIO_SIGNALPROCESSINGMODE_DEFAULT = "{C18E2F7E-933D-4965-B7D1-1EEF228D2AF3}"

# =============================================================================
# End Registry Constants
# =============================================================================

function Start-ProcessWithTimeout {
    <#
    .SYNOPSIS
        Starts a process, waits with timeout, and throws on timeout.
    .DESCRIPTION
        Used for external tools that can occasionally stall (e.g. regedit.exe).
        On timeout, the process is killed and a descriptive exception is thrown.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter(Mandatory)]
        [string]$ArgumentList,

        [Parameter(Mandatory)]
        [int]$TimeoutMs,

        [Parameter(Mandatory)]
        [string]$Context,

        [System.Diagnostics.ProcessWindowStyle]$WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    )

    $procName = [System.IO.Path]::GetFileName($FilePath)
    $proc = $null

    try {
        $proc = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru -WindowStyle $WindowStyle
    } catch {
        throw "[PROC] Failed to start '$procName' while ${Context}: $($_.Exception.Message)"
    }

    if (-not $proc) {
        throw "[PROC] Failed to start '$procName' while ${Context}: process handle was null"
    }

    $completed = $false
    try {
        $completed = $proc.WaitForExit($TimeoutMs)
    } catch {
        throw "[PROC] Failed while waiting for '$procName' while ${Context}: $($_.Exception.Message)"
    }

    if (-not $completed) {
        try { $proc.Kill() } catch { Write-Verbose "Failed to kill process '$procName': $_" }
        $timeoutSeconds = [Math]::Round(($TimeoutMs / 1000.0), 1)
        throw "[PROC-TIMEOUT] '$procName' timed out after ${timeoutSeconds}s while $Context"
    }

    return $proc
}

function Ensure-EnhancementCapability {
    <#
    .SYNOPSIS
        Ensures Windows audio enhancement capability by registering required COM CLSIDs.

    .DESCRIPTION
        Windows audio enhancements (LEQ, Bass Boost, Virtual Surround, etc.) require
        COM class registrations pointing to WMALFXGFXDSP.dll. Some systems are missing
        these registrations, causing E-APO to fail and enhancements to not work.

        This function registers all required CLSIDs:
        - LEQ APO: {62dc1a93-ae24-464c-a43e-452f824c4250}
        - OTHER_APO: {637c490d-eee3-4c0a-973f-371958802da2}
        - WM_GFX_APO: {13AB3EBD-137E-4903-9D89-60BE8277FD17}
        - WM_LFX_APO: {C9453E73-8C5C-4463-9984-AF8BAB2F5447}

        All point to the same Microsoft system DLL present on all Windows 10/11 systems.
        Safe to call multiple times - only creates missing registrations.

    .OUTPUTS
        [bool] True if all enhancements are now available

    .NOTES
        Requires administrator privileges.
    #>

    $dllPath = $script:ENHANCEMENT_DLL

    # All CLSIDs that need to be registered for full enhancement support across all APO modes
    $clsids = @(
        @{ Guid = $script:LEQ_APO_GUID; Name = "LEQ" },
        @{ Guid = $script:OTHER_APO_GUID; Name = "OTHER_APO" },
        @{ Guid = $script:WM_GFX_APO_GUID; Name = "WM_GFX_APO" },
        @{ Guid = $script:WM_LFX_APO_GUID; Name = "WM_LFX_APO" }
    )

    # Verify the system DLL exists (should always be true on Windows 10/11)
    if (-not (Test-Path $dllPath)) {
        Write-Warning "Enhancement DLL not found: $dllPath"
        return $false
    }

    $allSuccess = $true

    foreach ($entry in $clsids) {
        $clsid = $entry.Guid
        $name = $entry.Name
        $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid"
        $inprocPath = "$clsidPath\InprocServer32"

        # Skip if already registered
        if (Test-Path $clsidPath) {
            Write-Verbose "$name CLSID already registered"
            continue
        }

        try {
            Write-Verbose "Registering $name CLSID..."

            New-Item -Path $inprocPath -Force | Out-Null
            Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value $dllPath
            Set-ItemProperty -Path $inprocPath -Name "ThreadingModel" -Value "Both"

            Write-Verbose "$name CLSID registered successfully"
        }
        catch {
            Write-Warning "Failed to register $name CLSID: $_"
            $allSuccess = $false
        }
    }

    return $allSuccess
}

function Clear-CompositeFxKeys {
    <#
    .SYNOPSIS
        Removes CompositeFX keys from a device's FxProperties to enable LFX/GFX APO slots.

    .DESCRIPTION
        Some Windows configurations (Atlas OS, IoT LTSC, certain driver versions) add
        CompositeFX keys (slots 13-20 under {d04e05a6}) that force the modern composite
        APO path. This can prevent LFX/GFX slots from being used.

        Deleting these keys allows Windows to fall back to the legacy LFX/GFX model,
        which is required for our standard LEQ installation flow.

        IMPORTANT: This function only targets the real CompositeFX slots (13-20) under
        the FX_PROPERTY_BASE GUID. It does NOT touch {d3993a3f} processing mode keys
        (slots 5/6/7 etc.) which declare supported audio processing modes. Deleting
        those breaks audio routing through the APO chain.

        Uses .reg file import via regedit.exe due to MMDevices registry ACL restrictions.

    .PARAMETER DeviceGuid
        The GUID of the audio render device (without braces, or with braces - both work).

    .EXAMPLE
        Clear-CompositeFxKeys -DeviceGuid "b5b9a97a-d9f2-45ba-8302-76ea3e9266c6"

    .OUTPUTS
        [bool] True if operation completed (keys deleted or didn't exist), False on error.

    .NOTES
        Safe to call even if CompositeFX keys don't exist - will be a no-op.
        Requires administrator privileges.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DeviceGuid
    )

    # Normalize GUID format (remove braces if present)
    $DeviceGuid = $DeviceGuid.Trim('{}')

    $fxPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{$DeviceGuid}\FxProperties"

    # Check if FxProperties exists
    if (-not (Test-Path -LiteralPath $fxPath)) {
        Write-Verbose "FxProperties not found for device $DeviceGuid - skipping CompositeFX clear"
        return $true
    }

    # Check for real CompositeFX keys: {d04e05a6},13 through ,20
    # Slots 13=CompositeSFX, 14=CompositeMFX, 15=CompositeEFX, 16-20=Offload/KeywordDetector variants
    $fxProps = Get-ItemProperty -LiteralPath $fxPath -ErrorAction SilentlyContinue
    $compositeFxSlots = 13..20
    $compositeFxKeys = @($fxProps.PSObject.Properties | Where-Object {
        $propName = $_.Name
        foreach ($slot in $compositeFxSlots) {
            if ($propName -eq "$($script:FX_PROPERTY_BASE),$slot") { return $true }
        }
        return $false
    })

    if ($compositeFxKeys.Count -eq 0) {
        Write-Verbose "No CompositeFX keys found for device $DeviceGuid - LFX/GFX should work"
        return $true
    }

    Write-Output "[COMPOSITE FX] Found $($compositeFxKeys.Count) CompositeFX keys on device $DeviceGuid"
    Write-Output "[COMPOSITE FX] Removing to enable LFX/GFX APO slots..."

    # Build .reg file to delete CompositeFX keys
    # Using .reg file because MMDevices has special ACLs that block standard PowerShell cmdlets
    $regKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{$DeviceGuid}\FxProperties"

    $regContent = "Windows Registry Editor Version 5.00`r`n`r`n"
    $regContent += "[$regKeyPath]`r`n"

    foreach ($key in $compositeFxKeys) {
        $keyName = $key.Name
        $regContent += "`"$keyName`"=-`r`n"
        Write-Verbose "  Will delete: $keyName"
    }

    # Write and import .reg file
    $regFile = Join-Path $env:TEMP "AIW_ClearCompositeFx_$([System.IO.Path]::GetRandomFileName()).reg"
    try {
        $regContent | Out-File -FilePath $regFile -Encoding ASCII -Force
        Write-Verbose "[COMPOSITE FX] Created temp file: $regFile"

        $proc = Start-Process -FilePath "$env:SystemRoot\regedit.exe" -ArgumentList '/s', $regFile -Verb RunAs -Wait -PassThru -WindowStyle Hidden

        if ($proc.ExitCode -eq 0) {
            Write-Output "[COMPOSITE FX] [OK] CompositeFX keys removed successfully"
            return $true
        } else {
            Write-Output "[COMPOSITE FX] [ERROR] Registry import failed with exit code: $($proc.ExitCode)"
            return $false
        }
    }
    catch {
        Write-Output "[COMPOSITE FX] [ERROR] Failed to remove CompositeFX keys: $_"
        return $false
    }
    finally {
        if (Test-Path $regFile) {
            Remove-Item $regFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function Repair-LEQClsids {
    <#
    .SYNOPSIS
        Checks and repairs system-wide CLSID registrations needed for audio enhancements.
    .DESCRIPTION
        Validates that all 4 required CLSIDs exist with correct InprocServer32 values
        pointing to WMALFXGFXDSP.dll. Fixes any that are missing or misconfigured.
    .OUTPUTS
        [hashtable] With keys: Missing (names detected), Fixed (names repaired), Failed (names that couldn't be repaired)
    #>

    $dllPath = $script:ENHANCEMENT_DLL
    $clsids = @(
        @{ Guid = $script:LEQ_APO_GUID;    Name = "LEQ APO" },
        @{ Guid = $script:OTHER_APO_GUID;   Name = "Enhancement APO" },
        @{ Guid = $script:WM_GFX_APO_GUID;  Name = "WM GFX APO" },
        @{ Guid = $script:WM_LFX_APO_GUID;  Name = "WM LFX APO" }
    )

    $result = @{
        Missing  = @()
        Fixed    = @()
        Failed   = @()
    }

    if (-not (Test-Path $dllPath)) {
        Write-Warning "Enhancement DLL not found: $dllPath"
        return $result
    }

    foreach ($entry in $clsids) {
        $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$($entry.Guid)"
        $inprocPath = "$clsidPath\InprocServer32"

        # Validate existence AND correct values (not just key existence)
        $needsFix = $false
        if (-not (Test-Path $inprocPath)) {
            $needsFix = $true
        } else {
            $existing = Get-ItemProperty -LiteralPath $inprocPath -ErrorAction SilentlyContinue
            $currentDefault = if ($existing) { $existing.'(Default)' } else { $null }
            $currentThreading = if ($existing) { $existing.ThreadingModel } else { $null }
            if ($currentDefault -ne $dllPath -or $currentThreading -ne 'Both') {
                $needsFix = $true
            }
        }

        if (-not $needsFix) { continue }

        $result.Missing += $entry.Name
        try {
            New-Item -Path $inprocPath -Force | Out-Null
            Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value $dllPath
            Set-ItemProperty -Path $inprocPath -Name "ThreadingModel" -Value "Both"
            $result.Fixed += $entry.Name
        } catch {
            $result.Failed += $entry.Name
        }
    }

    return $result
}

function Test-EapoInFxSlots {
    param(
        [Parameter(Mandatory)] $FxProperties
    )

    if (-not $FxProperties) { return $false }

    # Check all FX slots where E-APO might install
    foreach ($slot in @(',1', ',2', ',5', ',6')) {
        $propName = "$script:FX_PROPERTY_BASE$slot"
        if ($FxProperties.PSObject.Properties.Name -contains $propName) {
            $value = $FxProperties.$propName
            if ($value) {
                $valueStr = $value.ToString().ToUpper()
                if ($valueStr -match [regex]::Escape($script:EAPO_PRE_MIX_GUID.ToUpper()) -or
                    $valueStr -match [regex]::Escape($script:EAPO_POST_MIX_GUID.ToUpper())) {
                    return $true
                }
            }
        }
    }
    return $false
}

function Get-DeviceAudioFormat {
    param(
        [Parameter(Mandatory)] [string]$DeviceGuid
    )

    $result = @{
        Channels = 2
        BitDepth = 16
        SampleRate = 48000
    }

    try {
        $propsPath = Join-Path $script:REG_MMDEVICES_RENDER "$DeviceGuid\Properties"
        if (-not (Test-Path -LiteralPath $propsPath)) {
            return $result
        }

        $props = Get-ItemProperty -LiteralPath $propsPath -ErrorAction SilentlyContinue
        if (-not $props -or ($props.PSObject.Properties.Name -notcontains $script:PROP_DEVICE_FORMAT)) {
            return $result
        }

        $formatBytes = $props.$($script:PROP_DEVICE_FORMAT)
        # WAVEFORMATEXTENSIBLE structure (after 8-byte header):
        # Offset 8:  wFormatTag (2 bytes) - 0xFFFE = WAVE_FORMAT_EXTENSIBLE
        # Offset 10: nChannels (2 bytes)
        # Offset 12: nSamplesPerSec (4 bytes)
        # Offset 16: nAvgBytesPerSec (4 bytes)
        # Offset 20: nBlockAlign (2 bytes)
        # Offset 22: wBitsPerSample (2 bytes) - container bit depth
        # Offset 24: cbSize (2 bytes)
        # Offset 26: wValidBitsPerSample (2 bytes) - actual bit depth for extensible
        if (-not $formatBytes -or $formatBytes.Length -lt 24) {
            return $result
        }

        $result.Channels = [BitConverter]::ToInt16($formatBytes, 10)
        $result.SampleRate = [BitConverter]::ToInt32($formatBytes, 12)
        $result.BitDepth = [BitConverter]::ToInt16($formatBytes, 22)

        # For WAVE_FORMAT_EXTENSIBLE (0xFFFE), use wValidBitsPerSample if available
        # This handles devices with 32-bit container but 24-bit actual audio (e.g., Elgato 4K X)
        $formatTag = [BitConverter]::ToInt16($formatBytes, 8)
        if ($formatTag -eq -2 -and $formatBytes.Length -ge 28) {  # -2 = 0xFFFE as signed Int16
            $validBits = [BitConverter]::ToInt16($formatBytes, 26)
            if ($validBits -ge 8 -and $validBits -le 32) {
                $result.BitDepth = $validBits
            }
        }

        return $result
    } catch {
        Write-Verbose "Get-DeviceAudioFormat failed: $_"
        return $result
    }
}

function Get-AudioDeviceInfo {
    param(
        [switch]$IncludeInactive
    )

    $devices = @()
    $renderRoot = $script:REG_MMDEVICES_RENDER
    $deviceKeys = Get-ChildItem -LiteralPath $renderRoot -ErrorAction SilentlyContinue

    foreach ($deviceKey in $deviceKeys) {
        $guid = Split-Path -Leaf $deviceKey.PSPath

        # Read device state from registry
        # DeviceState: 1=Active/Ready, 2=Disabled, 4=NotPresent, 8=Unplugged
        # High bits (e.g. 0x10000000) indicate non-ready states — match on exact value
        $deviceInfo = Get-ItemProperty -LiteralPath $deviceKey.PSPath -ErrorAction SilentlyContinue
        $deviceState = if ($deviceInfo -and $deviceInfo.PSObject.Properties.Name -contains 'DeviceState') {
            $deviceInfo.DeviceState
        } else { 0 }

        # Skip disabled/inactive endpoints unless explicitly requested
        # Only DeviceState=1 means the endpoint is enabled in Windows Sound settings
        if (-not $IncludeInactive -and $deviceState -ne 1) { continue }

        $propsPath = Join-Path $deviceKey.PSPath 'Properties'
        $props = Get-ItemProperty -LiteralPath $propsPath -ErrorAction SilentlyContinue

        # Friendly name fallback chain
        $name = $null
        $interfaceName = $null

        if ($props) {
            try {
                # Use friendly name as primary (e.g., "Speakers", "Headset")
                if ($props.PSObject.Properties.Name -contains $script:PROP_DEVICE_FRIENDLY_NAME) {
                    $name = $props.$($script:PROP_DEVICE_FRIENDLY_NAME)
                }

                # Use device description as interface name (e.g., "Sound Blaster GC7", "Realtek HD Audio")
                if ($props.PSObject.Properties.Name -contains $script:PROP_DEVICE_DESC) {
                    $interfaceName = $props.$($script:PROP_DEVICE_DESC)
                }

                # Fallback: If no friendly name, use description as the name
                if (-not $name -and $interfaceName) {
                    $name = $interfaceName
                    $interfaceName = $null
                }
            } catch {
                Write-Verbose "Failed to read device name for $guid : $_"
            }
        }

        if (-not $name) { continue }

        $fxPath = Join-Path $deviceKey.PSPath 'FxProperties'
        $supportsEnhancement = $false
        $hasLfxGfx = $false            # Does the device have LFX/GFX APO slots?
        $leqConfigured = $false        # NEW: Does the LEQ registry key exist?
        $loudnessEnabled = $false       # Existing: Is LEQ turned ON?
        $releaseTime = 4
        $hasReleaseTimeKey = $false
        $hasCompositeFx = $false        # Does the device have CompositeFX keys blocking LFX/GFX?
        $eapoStatus = "Not Installed"  # Initialize before conditional block
        $eapoChildBroken = $false      # Initialize before conditional block

        if (Test-Path -LiteralPath $fxPath) {
            $fxProps = Get-ItemProperty -LiteralPath $fxPath -ErrorAction SilentlyContinue
            if ($fxProps) {
                $supportsEnhancement = $fxProps.PSObject.Properties.Name -contains $script:PROP_ENHANCEMENT_TAB

                # Check if FxProperties has any registry values (not just the key)
                # Virtual devices without E-APO have an empty FxProperties key (0 values)
                # Hardware devices and E-APO configured devices always have values
                $fxItem = Get-Item -LiteralPath $fxPath -ErrorAction SilentlyContinue
                $hasLfxGfx = $fxItem -and ($fxItem.GetValueNames().Count -gt 0)

                # Check for CompositeFX keys ({d04e05a6},13-20) that block LFX/GFX
                $compositeFxSlots = 13..20
                foreach ($slot in $compositeFxSlots) {
                    if ($fxProps.PSObject.Properties.Name -contains "$($script:FX_PROPERTY_BASE),$slot") {
                        $hasCompositeFx = $true
                        break
                    }
                }

                # leqConfigured — enhancement tab OR toggle key existence
                $enhancementTabKey = '{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},3'
                $enhancementTabGuid = '{5860E1C5-F95C-4a7a-8EC8-8AEF24F379A1}'
                $ourLeqKey = '{fc52a749-4be9-4510-896e-966ba6525980},3'

                $hasEnhancementTab = $false
                if ($fxProps.PSObject.Properties.Name -contains $enhancementTabKey) {
                    $propValue = $fxProps.PSObject.Properties | Where-Object { $_.Name -eq $enhancementTabKey } | Select-Object -ExpandProperty Value
                    $hasEnhancementTab = $propValue -eq $enhancementTabGuid
                }
                $hasToggleKey = $fxProps.PSObject.Properties.Name -contains $ourLeqKey
                $leqConfigured = $hasEnhancementTab -or $hasToggleKey

                # loudnessEnabled (ON/OFF state)
                $loudnessEnabled = $false
                if ($hasToggleKey) {
                    $flag = $fxProps.PSObject.Properties | Where-Object { $_.Name -eq $ourLeqKey } | Select-Object -ExpandProperty Value
                    if ($flag -and $flag.Length -ge 10 -and $flag[8] -eq 255 -and $flag[9] -eq 255) {
                        $loudnessEnabled = $true
                        Write-Verbose "LEQ enabled detected in our key: $ourLeqKey"
                    }
                }

                # Get release time from FxProperties
                $releaseTime = 4  # Default
                $hasReleaseTimeKey = $false

                $rtBase = '{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0}'
                $rtKey = "$rtBase,3"
                if ($fxProps.PSObject.Properties.Name -contains $rtKey) {
                    try {
                        $rtValue = $fxProps.PSObject.Properties | Where-Object { $_.Name -eq $rtKey } | Select-Object -ExpandProperty Value
                        if ($rtValue -and $rtValue.Length -ge 9) {
                            $rtInt = [int]$rtValue[8]
                            if ($rtInt -ge 2 -and $rtInt -le 7) {
                                $releaseTime = $rtInt
                                $hasReleaseTimeKey = $true
                                Write-Verbose "Release time $rtInt detected in FxProperties key: $rtKey"
                            }
                        }
                    } catch {
                        Write-Verbose "Failed to read RT key $rtKey : $_"
                    }
                }

                # Check for Equalizer APO (E-APO) presence
                $eapoStatus = "Not Installed"

                # PRIMARY CHECK: E-APO APO GUIDs in FX slots (most reliable)
                $fxSlotsHaveEAPO = Test-EapoInFxSlots -FxProperties $fxProps

                # SECONDARY CHECK: E-APO's own registry key
                $childApoPath = "$script:EAPO_CHILD_APO_PATH\$guid"
                $childApoExists = Test-Path -LiteralPath $childApoPath

                # TERTIARY CHECK: E-APO DLL exists
                $eapoFilesExist = Test-Path "$env:ProgramFiles\EqualizerAPO\EqualizerAPO.dll"

                # Determine status - FX slots are authoritative
                $eapoStatus = switch ($true) {
                    ($fxSlotsHaveEAPO) { "Active" }
                    ($childApoExists -and $eapoFilesExist) { "Installed (not active on device)" }
                    ($childApoExists -and -not $eapoFilesExist) { "Not Installed" }  # Orphaned registry
                    default { "Not Installed" }
                }

                Write-Verbose "E-APO Status for $guid : FxSlots=$fxSlotsHaveEAPO, ChildAPO=$childApoExists, Files=$eapoFilesExist, Status=$eapoStatus"

                # Check if E-APO has LEQ in its Child APO chain
                # Only flag as broken when LEQ was previously installed (leqConfigured)
                # but got displaced from the chain. If LEQ was never installed, this is
                # just the normal pre-install state - not a broken chain.
                $eapoChildBroken = $false
                if ($eapoStatus -eq 'Active' -and $leqConfigured) {
                    $childApoPath = "HKLM:\SOFTWARE\EqualizerAPO\Child APOs\$($deviceKey.PSChildName)"
                    if (Test-Path $childApoPath) {
                        $childApo = Get-ItemProperty -LiteralPath $childApoPath -ErrorAction SilentlyContinue
                        $preMixChild = $childApo.PreMixChild
                        # LEQ APO GUID - must be in chain for LEQ toggle to work
                        $leqApoGuid = '{62dc1a93-ae24-464c-a43e-452f824c4250}'
                        # Broken if PreMixChild is empty OR doesn't contain LEQ APO
                        if ([string]::IsNullOrEmpty($preMixChild) -or $preMixChild -ne $leqApoGuid) {
                            $eapoChildBroken = $true
                        }
                    } else {
                        # No Child APO entry at all for this device
                        $eapoChildBroken = $true
                    }
                }
            }
        }

        $audioFormat = @{ Channels = 2; BitDepth = 16; SampleRate = 48000 }
        if ($deviceState -eq 1) {
            $audioFormat = Get-DeviceAudioFormat -DeviceGuid $guid
        }

        $devices += [PSCustomObject]@{
            Name                = $name
            InterfaceName       = $interfaceName
            Guid                = $guid
            State               = $deviceState
            IsActive            = ($deviceState -eq 1)
            SupportsEnhancement = $supportsEnhancement
            HasLfxGfx           = $hasLfxGfx
            HasCompositeFx      = $hasCompositeFx
            LeqConfigured       = $leqConfigured        # NEW PROPERTY
            LoudnessEnabled     = $loudnessEnabled
            ReleaseTime         = $releaseTime
            EapoStatus          = $eapoStatus
            EapoChildBroken     = $eapoChildBroken
            Channels            = $audioFormat.Channels
            BitDepth            = $audioFormat.BitDepth
            SampleRate          = $audioFormat.SampleRate
            HasReleaseTimeKey   = $hasReleaseTimeKey
        }
    }

    return $devices
}

function Get-EapoStatus {
    param(
        [Parameter(Mandatory)] [string]$DeviceGuid
    )

    # Get FxProperties
    $fxPath = Join-Path $script:REG_MMDEVICES_RENDER "$DeviceGuid\FxProperties"
    $fxProps = $null
    if (Test-Path -LiteralPath $fxPath) {
        $fxProps = Get-ItemProperty -LiteralPath $fxPath -ErrorAction SilentlyContinue
    }

    # PRIMARY: Check FX slots for E-APO GUIDs
    $fxSlotsHaveEAPO = Test-EapoInFxSlots -FxProperties $fxProps

    # Return true only if E-APO is actually in the FX slots
    return $fxSlotsHaveEAPO
}

function Install-LEQRegistry {
    param(
        [Parameter(Mandatory)] [PSCustomObject]$Device,
        [Parameter(Mandatory)] [ValidateRange(2,7)] [int]$ReleaseTime,
        [switch]$Force
    )

    Write-Output "=== INSTALL-LEQ ==="
    Write-Output "Device: $($Device.Name) ($($Device.Guid))"
    if ($Force) { Write-Output "[INSTALL] Force/Clean Install mode enabled" }

    if ($null -eq $Device -or [string]::IsNullOrWhiteSpace($Device.Guid)) {
        throw "Invalid device parameter"
    }

    # === STEP 1: Ensure system-wide enhancement capability (CLSID registration) ===
    if (-not (Ensure-EnhancementCapability)) {
        Write-Warning "Could not ensure enhancement capability, proceeding anyway..."
    }

    # === STEP 2: Clear CompositeFX keys to enable LFX/GFX path ===
    # This is critical for Atlas OS / IoT LTSC / certain driver configurations
    if (-not (Clear-CompositeFxKeys -DeviceGuid $Device.Guid)) {
        Write-Warning "Could not clear CompositeFX keys, LFX/GFX may not work"
    }

    $fxKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\$($Device.Guid)\FxProperties"

    # Check FxProperties key exists
    if (-not (Test-Path -LiteralPath $fxKeyPath)) {
        Write-Output "[INSTALL] FxProperties doesn't exist - device may not support enhancements"
        throw "Device does not support audio enhancements (FxProperties missing)"
    }

    # === STEP 2b: Snapshot existing processing mode keys ===
    # {d3993a3f},5 = PKEY_SFX_ProcessingModes_Supported_For_Streaming
    # {d3993a3f},6 = PKEY_MFX_ProcessingModes_Supported_For_Streaming
    # These must be preserved - without them the audio engine won't route through APOs.
    $fxPropsSnapshot = Get-ItemProperty -LiteralPath $fxKeyPath -ErrorAction SilentlyContinue
    $existingProcessingModes = @{}
    if ($null -ne $fxPropsSnapshot) {
        foreach ($prop in $fxPropsSnapshot.PSObject.Properties) {
            if ($prop.Name -match [regex]::Escape($script:PROCESSING_MODES_PROPERTY_KEY)) {
                $existingProcessingModes[$prop.Name] = $prop.Value
                Write-Verbose "[INSTALL] Snapshotted processing mode key: $($prop.Name)"
            }
        }
    }
    if ($existingProcessingModes.Count -gt 0) {
        Write-Output "[INSTALL] Preserving $($existingProcessingModes.Count) processing mode keys"
    } else {
        Write-Output "[INSTALL] No existing processing mode keys - will add defaults for SFX/MFX"
    }

    # Check if already has LEQ configured
    $fxProps = Get-ItemProperty -LiteralPath $fxKeyPath -ErrorAction SilentlyContinue

    # Check if Enhancement tab is already properly configured (the REAL indicator)
    $hasEnhancementTab = $false
    if ($null -ne $fxProps) {
        $enhancementSlotKey = "$($script:FX_PROPERTY_BASE)$($script:FX_SLOT_ENHANCEMENT)"
        if ($fxProps.PSObject.Properties.Name -contains $enhancementSlotKey) {
            $existingValue = $fxProps.PSObject.Properties | Where-Object { $_.Name -eq $enhancementSlotKey } | Select-Object -ExpandProperty Value
            if ($existingValue -eq $script:ENHANCEMENT_TAB_GUID) {
                $hasEnhancementTab = $true
            }
        }
    }

    if ($hasEnhancementTab -and -not $Force) {
        Write-Output "[INSTALL] Enhancement tab already configured - use Set-LEQRegistry to toggle (or use Clean Install)"
        return $false
    }

    if ($Force -and $hasEnhancementTab) {
        Write-Output "[INSTALL] Force mode - reinstalling FX properties over existing"
    } elseif (-not $hasEnhancementTab) {
        Write-Output "[INSTALL] Enhancement tab not found - proceeding with full install"
    }

    # Convert registry path for .reg file
    $regKeyPath = $fxKeyPath -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\'

    # Build .reg file with APO associations AND LEQ enable keys
    $regContent = "Windows Registry Editor Version 5.00`r`n`r`n"
    $regContent += "[$regKeyPath]`r`n"

    # FX APO associations (required for LEQ to work)
    $regContent += "`"$($script:FX_PROPERTY_BASE)$($script:FX_SLOT_LFX)`"=`"$($script:LEQ_APO_GUID)`"`r`n"
    $regContent += "`"$($script:FX_PROPERTY_BASE)$($script:FX_SLOT_GFX)`"=`"$($script:OTHER_APO_GUID)`"`r`n"
    $regContent += "`"$($script:FX_PROPERTY_BASE)$($script:FX_SLOT_ENHANCEMENT)`"=`"$($script:ENHANCEMENT_TAB_GUID)`"`r`n"
    $regContent += "`"$($script:FX_PROPERTY_BASE)$($script:FX_SLOT_SFX)`"=`"$($script:LEQ_APO_GUID)`"`r`n"
    $regContent += "`"$($script:FX_PROPERTY_BASE)$($script:FX_SLOT_MFX)`"=`"$($script:OTHER_APO_GUID)`"`r`n"

    # LEQ enable keys (enabled by default on install)
    $enabledHex = "0b,00,00,00,01,00,00,00,ff,ff,00,00"
    $regContent += "`"{fc52a749-4be9-4510-896e-966ba6525980},3`"=hex:$enabledHex`r`n"
    $regContent += "`"{fc52a749-4be9-4510-896e-966ba6525980},1599`"=hex:$enabledHex`r`n"

    # Release time keys
    $releaseTimeHex = "03,00,00,00,01,00,00,00,$($ReleaseTime.ToString('x2')),00,00,00"
    $regContent += "`"{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},3`"=hex:$releaseTimeHex`r`n"
    $regContent += "`"{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},1599`"=hex:$releaseTimeHex`r`n"

    # Processing mode keys - preserve existing or add defaults
    # These tell the audio engine which processing modes the SFX/MFX APOs support.
    # Without them, LEQ loads but the audio engine won't route audio through it.
    if ($existingProcessingModes.Count -gt 0) {
        # Restore snapshotted keys
        foreach ($entry in $existingProcessingModes.GetEnumerator()) {
            $keyName = $entry.Key
            $values = @($entry.Value)
            # REG_MULTI_SZ: encode as hex(7) - UTF-16LE strings with double null terminator
            $hexParts = @()
            foreach ($v in $values) {
                $bytes = [System.Text.Encoding]::Unicode.GetBytes($v)
                $hexParts += ($bytes | ForEach-Object { $_.ToString('x2') }) -join ','
                $hexParts += '00,00'  # null terminator for this string
            }
            $hexParts += '00,00'  # final null terminator for MULTI_SZ
            $hexStr = $hexParts -join ','
            $regContent += "`"$keyName`"=hex(7):$hexStr`r`n"
        }
        Write-Verbose "[INSTALL] Restored $($existingProcessingModes.Count) processing mode keys"
    } else {
        # Add default processing mode keys for SFX and MFX
        # AUDIO_SIGNALPROCESSINGMODE_DEFAULT encoded as REG_MULTI_SZ hex(7)
        $defaultModeBytes = [System.Text.Encoding]::Unicode.GetBytes($script:AUDIO_SIGNALPROCESSINGMODE_DEFAULT)
        $defaultModeHex = (($defaultModeBytes | ForEach-Object { $_.ToString('x2') }) -join ',') + ',00,00,00,00'
        $sfxKey = "$($script:PROCESSING_MODES_PROPERTY_KEY),5"
        $mfxKey = "$($script:PROCESSING_MODES_PROPERTY_KEY),6"
        $regContent += "`"$sfxKey`"=hex(7):$defaultModeHex`r`n"
        $regContent += "`"$mfxKey`"=hex(7):$defaultModeHex`r`n"
        Write-Verbose "[INSTALL] Added default SFX/MFX processing mode keys"
    }

    # UI state notification key in Properties (tells Windows Sound applet LEQ is ON)
    $propsKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\$($Device.Guid)\Properties"
    $propsRegKeyPath = $propsKeyPath -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\'
    $regContent += "`r`n[$propsRegKeyPath]`r`n"
    $regContent += "`"{1e94c58f-3e40-4ddb-b10c-a86d8b870a31},2`"=hex:02,00,00,00,01,00,00,00,fb,02`r`n"

    # User subkey - some devices read LEQ/RT state from here
    $userKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\$($Device.Guid)\FxProperties\{b13412ee-07af-4c57-b08b-e327f8db085b}\User"
    if (Test-Path -LiteralPath $userKeyPath) {
        $userRegKeyPath = $userKeyPath -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\'
        $regContent += "`r`n[$userRegKeyPath]`r`n"
        $regContent += "`"{fc52a749-4be9-4510-896e-966ba6525980},3`"=hex:$enabledHex`r`n"
        $regContent += "`"{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},3`"=hex:$releaseTimeHex`r`n"
        Write-Verbose "[INSTALL] Writing LEQ enable + RT to User subkey"
    }

    # Write and import .reg file
    $regFile = Join-Path $env:TEMP "AIW_LEQ_Install_$([System.IO.Path]::GetRandomFileName()).reg"
    try {
        $regContent | Out-File -FilePath $regFile -Encoding ASCII -Force
        Write-Output "[INSTALL] Created: $regFile"

        $proc = Start-Process -FilePath "$env:SystemRoot\regedit.exe" -ArgumentList '/s', $regFile -Verb RunAs -Wait -PassThru -WindowStyle Hidden

        if ($proc.ExitCode -eq 0) {
            Write-Output "[INSTALL] [OK] Registry import successful"

            Write-Output "[INSTALL] Restarting audio service..."
            try {
                Restart-Service audiosrv -Force -ErrorAction Stop
                Write-Output "[INSTALL] [OK] Audio service restarted"
            } catch {
                Write-Output "[INSTALL] [WARNING] Audio service restart failed: $($_.Exception.Message)"
            }

            Write-Output "[INSTALL] [OK] LEQ installed successfully"
            return $true
        } else {
            Write-Output "[INSTALL] [ERROR] Registry import failed: Exit code $($proc.ExitCode)"
            return $false
        }
    } finally {
        if (Test-Path $regFile) {
            Remove-Item $regFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function Set-ReleaseTime {
    param(
        [Parameter(Mandatory)] [string]$DeviceGuid,
        [Parameter(Mandatory)] [ValidateRange(2,7)] [int]$ReleaseTime
    )

    Write-Output "[RT] Setting Release Time to $ReleaseTime for device $DeviceGuid"

    $fxKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\$DeviceGuid\FxProperties"

    if (-not (Test-Path -LiteralPath $fxKeyPath)) {
        Write-Output "[RT] [ERROR] FxProperties not found"
        return $false
    }

    # Convert registry path for .reg file
    $regKeyPath = $fxKeyPath -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\'

    # Build .reg file with release time keys
    $regContent = "Windows Registry Editor Version 5.00`r`n`r`n"
    $regContent += "[$regKeyPath]`r`n"

    # Release time hex format
    $releaseTimeHex = "03,00,00,00,01,00,00,00,$($ReleaseTime.ToString('x2')),00,00,00"

    # Write to RT keys in FxProperties
    $regContent += "`"{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},3`"=hex:$releaseTimeHex`r`n"
    $regContent += "`"{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},1599`"=hex:$releaseTimeHex`r`n"

    # Also write to User subkey for devices that need it
    $userKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\$DeviceGuid\FxProperties\{b13412ee-07af-4c57-b08b-e327f8db085b}\User"
    if (Test-Path -LiteralPath $userKeyPath) {
        $userRegKeyPath = $userKeyPath -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\'
        $regContent += "`r`n[$userRegKeyPath]`r`n"
        $regContent += "`"{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},3`"=hex:$releaseTimeHex`r`n"
    }

    # Write and import .reg file
    $regFile = Join-Path $env:TEMP "AIW_RT_$([System.IO.Path]::GetRandomFileName()).reg"
    try {
        $regContent | Out-File -FilePath $regFile -Encoding ASCII -Force
        Write-Output "[RT] Created: $regFile"

        $proc = Start-Process -FilePath "$env:SystemRoot\regedit.exe" -ArgumentList '/s', $regFile -Verb RunAs -Wait -PassThru -WindowStyle Hidden

        if ($proc.ExitCode -ne 0) {
            Write-Output "[RT] [ERROR] Registry import failed"
            return $false
        }

        # Verify the write by reading back the primary RT key
        Start-Sleep -Milliseconds 200
        $rtVerifyKey = "{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},3"
        try {
            $currentValue = (Get-ItemProperty -LiteralPath $fxKeyPath -Name $rtVerifyKey -ErrorAction Stop).$rtVerifyKey
            if ($currentValue -is [byte[]] -and $currentValue.Length -ge 9 -and $currentValue[8] -eq $ReleaseTime) {
                Write-Output "[RT] [VERIFY] Release Time byte[8] = $($currentValue[8]) [MATCH]"
                Write-Output "[RT] [OK] Release Time set to $ReleaseTime"
                return $true
            } else {
                $actual = if ($currentValue -is [byte[]] -and $currentValue.Length -ge 9) { $currentValue[8] } else { "N/A" }
                Write-Output "[RT] [VERIFY] Release Time byte[8] = $actual [WRONG] (expected: $ReleaseTime)"
                return $false
            }
        } catch {
            Write-Output "[RT] [VERIFY] Failed to read back RT key: $($_.Exception.Message)"
            return $false
        }
    } finally {
        if (Test-Path $regFile) {
            Remove-Item $regFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function Set-LEQRegistry {
    param(
        [Parameter(Mandatory)] [PSCustomObject]$Device,
        [Parameter(Mandatory)] [bool]$Enabled,
        [Parameter(Mandatory)] [ValidateRange(2,7)] [int]$ReleaseTime
    )

    # DIAGNOSTIC: Log what we received
    Write-Output "=== SET-LEQ DIAGNOSTIC ==="
    Write-Output "Device type: $($Device.GetType().FullName)"
    Write-Output "Device.Guid: $($Device.Guid)"
    Write-Output "[PS RESULT] Device.State: $($Device.State) | IsActive: $($Device.IsActive)"
    Write-Output "Enabled: $Enabled"
    Write-Output "ReleaseTime: $ReleaseTime"

    if ($null -eq $Device) {
        Write-Output "ERROR: Device is null!"
        throw "Device parameter is null"
    }

    if ([string]::IsNullOrWhiteSpace($Device.Guid)) {
        Write-Output "ERROR: Device.Guid is null or empty!"
        throw "Device.Guid is missing"
    }

    $fxKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\$($Device.Guid)\FxProperties"
    Write-Output "FxKeyPath: $fxKeyPath"
    Write-Output "Path exists: $(Test-Path -LiteralPath $fxKeyPath)"
    Write-Output "=== END DIAGNOSTIC ==="

    if (-not (Test-Path -LiteralPath $fxKeyPath)) {
        throw "Device does not support audio enhancements (FxProperties missing)"
    }

    # Grant write access to the FxProperties key
    try {
        Write-Output "[ACL] Taking ownership of registry key..."

        $acl = Get-Acl -LiteralPath $fxKeyPath
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

        # Add Full Control for current user
        $rule = New-Object System.Security.AccessControl.RegistryAccessRule(
            $currentUser,
            [System.Security.AccessControl.RegistryRights]::FullControl,
            [System.Security.AccessControl.InheritanceFlags]::ContainerInherit,
            [System.Security.AccessControl.PropagationFlags]::None,
            [System.Security.AccessControl.AccessControlType]::Allow
        )

        $acl.AddAccessRule($rule)
        Set-Acl -LiteralPath $fxKeyPath -AclObject $acl

        Write-Output "[ACL] [OK] Granted write access"
    } catch {
        Write-Output "[ACL] [ERROR] Failed to modify permissions: $_"
        Write-Output "[ACL] Attempting writes anyway..."
    }

    $fxProps = Get-ItemProperty -LiteralPath $fxKeyPath -ErrorAction SilentlyContinue

    # LEQ enable keys - include 1599 variants
    $ourLeqKey3 = '{fc52a749-4be9-4510-896e-966ba6525980},3'
    $ourLeqKey1599 = '{fc52a749-4be9-4510-896e-966ba6525980},1599'
    $rtBase = '{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0}'

    # Find existing LEQ and RT keys
    $existingLeqKeys = @()
    $existingRtKeys = @()

    if ($null -ne $fxProps) {
        $existingRtKeys = @($fxProps.PSObject.Properties | Where-Object { $_.Name -match [regex]::Escape($rtBase) } | Select-Object -ExpandProperty Name)
    }

    $existingLeqKeys = @($ourLeqKey3, $ourLeqKey1599)

    Write-Output "[DEBUG] Will write to $($existingLeqKeys.Count) LEQ keys total"

    # If no RT keys exist yet, use the standard variations (match what Install writes)
    if ($existingRtKeys.Count -eq 0) {
        $existingRtKeys = @(
            "$rtBase,3"
            "$rtBase,1599"
        )
    }

    # === DIAGNOSTIC: Show current registry state BEFORE write ===
    Write-Output "`n[VERIFY] === BEFORE WRITE ==="
    Write-Output "[VERIFY] Keys we plan to write:"
    foreach ($key in $existingLeqKeys) {
        try {
            $currentValue = (Get-ItemProperty -LiteralPath $fxKeyPath -Name $key -ErrorAction Stop).$key
            $hexStr = ($currentValue | ForEach-Object { "{0:x2}" -f $_ }) -join ','
            Write-Output "[VERIFY]   LEQ $key = $hexStr"
        } catch {
            Write-Output "[VERIFY]   LEQ $key = (does not exist)"
        }
    }
    foreach ($key in $existingRtKeys) {
        try {
            $currentValue = (Get-ItemProperty -LiteralPath $fxKeyPath -Name $key -ErrorAction Stop).$key
            $hexStr = ($currentValue | ForEach-Object { "{0:x2}" -f $_ }) -join ','
            Write-Output "[VERIFY]   RT  $key = $hexStr"
        } catch {
            Write-Output "[VERIFY]   RT  $key = (does not exist)"
        }
    }
    Write-Output "[VERIFY] === END BEFORE ==="

    # Convert registry path from PowerShell format to .reg format
    $regKeyPath = $fxKeyPath -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\'

    # Build .reg file content
    $regContent = "Windows Registry Editor Version 5.00`r`n`r`n"
    $regContent += "[$regKeyPath]`r`n"

    # Convert to hex strings for .reg file format
    $enabledHex = if ($Enabled) {
        "0b,00,00,00,01,00,00,00,ff,ff,00,00"
    } else {
        "0b,00,00,00,01,00,00,00,00,00,00,00"
    }
    $releaseTimeHex = "03,00,00,00,01,00,00,00,$($ReleaseTime.ToString('x2')),00,00,00"

    # Add LEQ enable keys
    foreach ($key in $existingLeqKeys) {
        $regContent += "`"$key`"=hex:$enabledHex`r`n"
    }

    # Add Release Time keys
    foreach ($key in $existingRtKeys) {
        $regContent += "`"$key`"=hex:$releaseTimeHex`r`n"
    }

    # Also write to Properties key for Windows UI sync
    $propsKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\$($Device.Guid)\Properties"
    $propsRegKeyPath = $propsKeyPath -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\'

    $regContent += "`r`n[$propsRegKeyPath]`r`n"

    # UI state notification key
    $uiStateHex = if ($Enabled) {
        "02,00,00,00,01,00,00,00,fb,02"  # ON
    } else {
        "02,00,00,00,01,00,00,00,fb,01"  # OFF
    }
    $regContent += "`"{1e94c58f-3e40-4ddb-b10c-a86d8b870a31},2`"=hex:$uiStateHex`r`n"

    # Also write to User subkey for devices that need it
    $userKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\$($Device.Guid)\FxProperties\{b13412ee-07af-4c57-b08b-e327f8db085b}\User"
    if (Test-Path -LiteralPath $userKeyPath) {
        $userRegKeyPath = $userKeyPath -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\'
        $regContent += "`r`n[$userRegKeyPath]`r`n"
        $regContent += "`"{fc52a749-4be9-4510-896e-966ba6525980},3`"=hex:$enabledHex`r`n"
        $regContent += "`"{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},3`"=hex:$releaseTimeHex`r`n"
    }

    # Write .reg file to temp location
    $regFile = Join-Path $env:TEMP "AIW_LEQ_$([System.IO.Path]::GetRandomFileName()).reg"
    try {
        $regContent | Out-File -FilePath $regFile -Encoding ASCII -Force
        Write-Output "[REG] Created temp file: $regFile"

        # Import using regedit.exe with elevation
        $startParams = @{
            FilePath = "$env:SystemRoot\regedit.exe"
            ArgumentList = '/s', $regFile
            Verb = 'RunAs'
            Wait = $true
            WindowStyle = 'Hidden'
            PassThru = $true
        }

        Write-Output "[REG] Importing registry file..."
        $proc = Start-Process @startParams

            if ($proc.ExitCode -eq 0) {
            Write-Output "[REG] [OK] Registry import successful"

        } else {
            Write-Output "[REG] [ERROR] Registry import failed with exit code: $($proc.ExitCode)"
        }

    } finally {
        # Clean up temp file
        if (Test-Path $regFile) {
            Remove-Item $regFile -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Output "[REG] Registry writes completed"

    # === Verify registry state AFTER write ===
    Write-Output "`n[VERIFY] === AFTER WRITE ==="
    $verifyFailed = $false
    foreach ($key in $existingLeqKeys) {
        try {
            $currentValue = (Get-ItemProperty -LiteralPath $fxKeyPath -Name $key -ErrorAction Stop).$key
            $hexStr = ($currentValue | ForEach-Object { "{0:x2}" -f $_ }) -join ','
            $expectedHex = $enabledHex
            if ($hexStr -eq $expectedHex) {
                Write-Output "[VERIFY]   LEQ $key = $hexStr [MATCH]"
            } else {
                Write-Output "[VERIFY]   LEQ $key = $hexStr [WRONG] (expected: $expectedHex)"
                $verifyFailed = $true
            }
        } catch {
            Write-Output "[VERIFY]   LEQ $key = (STILL MISSING!)"
            $verifyFailed = $true
        }
    }
    foreach ($key in $existingRtKeys) {
        try {
            $currentValue = (Get-ItemProperty -LiteralPath $fxKeyPath -Name $key -ErrorAction Stop).$key
            $hexStr = ($currentValue | ForEach-Object { "{0:x2}" -f $_ }) -join ','
            $expectedHex = $releaseTimeHex
            if ($hexStr -eq $expectedHex) {
                Write-Output "[VERIFY]   RT  $key = $hexStr [MATCH]"
            } else {
                Write-Output "[VERIFY]   RT  $key = $hexStr [WRONG] (expected: $expectedHex)"
                $verifyFailed = $true
            }
        } catch {
            Write-Output "[VERIFY]   RT  $key = (STILL MISSING!)"
            $verifyFailed = $true
        }
    }
    Write-Output "[VERIFY] === END AFTER ===`n"

    if ($verifyFailed) {
        throw "Registry write verification failed — one or more values did not match after write"
    }

}
