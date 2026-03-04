// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading;
using System.Threading.Tasks;
using LEQControlPanel.Models;
using Microsoft.Win32;

namespace LEQControlPanel.Services;

internal sealed class AudioService
{
    private readonly InitialSessionState _initialState;
    private readonly bool _isInitialized;

    public AudioService()
    {
        try
        {
            _initialState = InitialSessionState.CreateDefault();
            _initialState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Unrestricted;

            _isInitialized = true;
        }
        catch (Exception)
        {
#if DEBUG
            Debug.WriteLine("[AUDIO SERVICE] Failed to initialize");
#endif
            throw;
        }
    }

    private class ConfiguredPowerShell : IDisposable
    {
        private readonly PowerShell _powerShell;
        private readonly Runspace _runspace;
        private bool _disposed;

        public ConfiguredPowerShell(PowerShell powerShell, Runspace runspace)
        {
            _powerShell = powerShell;
            _runspace = runspace;
        }

        public PowerShell PowerShell => _powerShell;

        public void Dispose()
        {
            if (!_disposed)
            {
                _powerShell?.Dispose();
                _runspace?.Dispose();
                _disposed = true;
            }
        }
    }

    private ConfiguredPowerShell CreateConfiguredPowerShell()
    {
        if (!_isInitialized)
        {
            throw new InvalidOperationException("PowerShell environment not initialized");
        }

        // Create runspace from configured template
        var runspace = RunspaceFactory.CreateRunspace(_initialState);
        runspace.Open();

        var ps = PowerShell.Create();
        ps.Runspace = runspace;

        // Load embedded scripts and dot-source to activate functions and variables
        var leqEngine = LoadEmbeddedScript("LEQControlPanel.LEQ-Engine.ps1");

        ps.AddScript(leqEngine);
        InvokeWithTimeout(ps);
        ps.Commands.Clear();

        return new ConfiguredPowerShell(ps, runspace);
    }

    /// <summary>
    /// Enumerates all audio devices via PowerShell and returns the PSObject matching
    /// the given device GUID, or null if not found.
    /// </summary>
    private static PSObject? FindDeviceByGuid(PowerShell ps, string deviceId)
    {
        ps.AddCommand("Get-AudioDeviceInfo");
        var devices = InvokeWithTimeout(ps);
        ps.Commands.Clear();

        foreach (PSObject device in devices)
        {
            if (device.Properties["Guid"]?.Value?.ToString() == deviceId)
                return device;
        }
        return null;
    }

    private static readonly TimeSpan PsTimeout = TimeSpan.FromSeconds(10);

    /// <summary>
    /// Invokes a PowerShell pipeline with a 10-second timeout. If the timeout
    /// fires, the pipeline is stopped and a TimeoutException is thrown.
    /// </summary>
    private static PSDataCollection<PSObject> InvokeWithTimeout(PowerShell ps)
    {
        var asyncResult = ps.BeginInvoke();
        if (asyncResult.AsyncWaitHandle.WaitOne(PsTimeout))
        {
            return ps.EndInvoke(asyncResult);
        }

        // Timeout — stop the pipeline and throw
        ps.Stop();
        throw new TimeoutException("PowerShell command timed out after 10 seconds");
    }

    public async Task<IReadOnlyList<AudioDevice>> GetDevicesAsync(bool includeInactive = false)
    {
        return await Task.Run(() =>
        {
            var devices = new List<AudioDevice>();

            try
            {
                using (var psWrapper = CreateConfiguredPowerShell())
                {
                    var ps = psWrapper.PowerShell;

                    // Now just call the function - no script concatenation needed
                    ps.AddCommand("Get-AudioDeviceInfo")
                      .AddParameter("IncludeInactive", includeInactive);

                    var results = InvokeWithTimeout(ps);

                    // Check for PowerShell errors
                    if (ps.Streams.Error.Count > 0)
                    {
                        Debug.WriteLine($"[DEVICE] PowerShell error count: {ps.Streams.Error.Count}");

                        foreach (var error in ps.Streams.Error)
                        {
                            Debug.WriteLine("=== PowerShell Error ===");
                            Debug.WriteLine($"Error: {error}");
                            Debug.WriteLine($"Exception: {error.Exception?.Message}");
                            Debug.WriteLine($"Script Stack Trace: {error.ScriptStackTrace}");
                            Debug.WriteLine($"Category: {error.CategoryInfo}");
                            Debug.WriteLine("======================");
                        }

                        var allErrors = string.Join("\n", ps.Streams.Error.Select(e => e.ToString()));
                        throw new Exception($"PowerShell execution failed:\n{allErrors}");
                    }
                    // If we get here, error count was 0, continue processing results

                    Debug.WriteLine($"[DEVICE] PowerShell completed successfully, processing {results.Count} results");

                    foreach (PSObject result in results)
                    {
                        try
                        {
                            if (result is null)
                            {
                                continue;
                            }

                            var guid = result.Properties["Guid"]?.Value?.ToString();
                            var name = result.Properties["Name"]?.Value?.ToString();
                            var interfaceName = result.Properties["InterfaceName"]?.Value?.ToString() ?? string.Empty;
#if DEBUG
                            Debug.WriteLine($"[DEVICE] Name: {name}, InterfaceName: '{interfaceName}'");
#endif
                            var releaseTimeValue = Convert.ToInt32(result.Properties["ReleaseTime"]?.Value ?? 4);
                            var supportsEnhancement = ParseBoolean(result.Properties["SupportsEnhancement"]?.Value);
                            var hasLfxGfx = ParseBoolean(result.Properties["HasLfxGfx"]?.Value);
                            var leqConfigured = ParseBoolean(result.Properties["LeqConfigured"]?.Value);
                            var loudnessEnabled = ParseBoolean(result.Properties["LoudnessEnabled"]?.Value);
                            var eapoStatus = result.Properties["EapoStatus"]?.Value?.ToString() ?? "Missing";
                            var eapoChildBroken = ParseBoolean(result.Properties["EapoChildBroken"]?.Value);
                            var channels = Convert.ToInt32(result.Properties["Channels"]?.Value ?? 2);
                            var bitDepth = Convert.ToInt32(result.Properties["BitDepth"]?.Value ?? 16);
                            var sampleRate = Convert.ToInt32(result.Properties["SampleRate"]?.Value ?? 48000);
                            var hasReleaseTimeKey = ParseBoolean(result.Properties["HasReleaseTimeKey"]?.Value);
                            var hasCompositeFx = ParseBoolean(result.Properties["HasCompositeFx"]?.Value);

                            if (string.IsNullOrWhiteSpace(guid) && string.IsNullOrWhiteSpace(name))
                            {
                                continue;
                            }

                            var device = new AudioDevice
                            {
                                Guid = guid ?? string.Empty,
                                Name = name ?? string.Empty,
                                InterfaceName = interfaceName,
                                SupportsEnhancement = supportsEnhancement,
                                HasLfxGfx = hasLfxGfx,
                                HasCompositeFx = hasCompositeFx,
                                LeqConfigured = leqConfigured,
                                LoudnessEnabled = loudnessEnabled,
                                ReleaseTime = releaseTimeValue,
                                EapoStatus = eapoStatus,
                                EapoChildBroken = eapoChildBroken,
                                Channels = channels,
                                BitDepth = bitDepth,
                                SampleRate = sampleRate,
                                HasReleaseTimeKey = hasReleaseTimeKey
                            };
                            devices.Add(device);
                        }
                        catch (Exception)
                        {
#if DEBUG
                            Debug.WriteLine("AudioService: Error parsing device result");
#endif
                            // Skip this device and continue with others
                            continue;
                        }
                    }

                    if (devices.Count == 0)
                    {
#if DEBUG
                        Debug.WriteLine($"[DEVICE] Filtering left zero valid devices");
#endif
                        devices.Add(new AudioDevice
                        {
                            Guid = "no_valid_devices",
                            Name = "No Valid Devices Found",
                            ReleaseTime = 4,
                            SupportsEnhancement = false,
                            LeqConfigured = false,
                            LoudnessEnabled = false,
                            EapoStatus = "Missing",
                            EapoChildBroken = false
                        });
                    }

                    return devices;
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                Debug.WriteLine($"FULL DEVICE ERROR: {ex.ToString()}");
                Debug.WriteLine($"ERROR MESSAGE: {ex.Message}");
                Debug.WriteLine($"STACK TRACE: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Debug.WriteLine($"INNER EXCEPTION: {ex.InnerException.Message}");
                    Debug.WriteLine($"INNER STACK TRACE: {ex.InnerException.StackTrace}");
                }
#endif
                devices.Add(new AudioDevice { Guid = "error", Name = $"Error: {ex.Message}" });
                return devices;
            }
        });
    }

    public async Task<bool?> ToggleLeqAsync(string deviceId)
    {
        return await Task.Run(() =>
        {
            try
            {
                using (var psWrapper = CreateConfiguredPowerShell())
                {
                    var ps = psWrapper.PowerShell;

                    var targetDevice = FindDeviceByGuid(ps, deviceId);
                    if (targetDevice == null)
                    {
                        #if DEBUG
                        Debug.WriteLine($"[LEQ TOGGLE] Device not found: {deviceId}");
                        #endif
                        return (bool?)null;
                    }

                    // Get current state
                    bool currentEnabled = Convert.ToBoolean(targetDevice.Properties["LoudnessEnabled"]?.Value ?? false);
                    int releaseTime = Convert.ToInt32(targetDevice.Properties["ReleaseTime"]?.Value ?? 4);
                    bool newEnabled = !currentEnabled;

                    #if DEBUG
                    Debug.WriteLine($"[LEQ TOGGLE] Device: {deviceId}, Current: {currentEnabled}, New: {newEnabled}");
                    #endif

                    // Use AddScript instead of AddCommand for proper PSObject passing
                    var script = @"
                        param($dev, $ena, $rt)
                        Set-LEQRegistry -Device $dev -Enabled $ena -ReleaseTime $rt
                    ";

                    ps.AddScript(script)
                      .AddParameter("dev", targetDevice)
                      .AddParameter("ena", newEnabled)
                      .AddParameter("rt", releaseTime);

                    var results = InvokeWithTimeout(ps);

                    // Capture Write-Host output from Information stream
                    if (ps.Streams.Information.Count > 0)
                    {
                        #if DEBUG
                        Debug.WriteLine($"[LEQ TOGGLE] PowerShell Output:");
                        #endif
                        foreach (var info in ps.Streams.Information)
                        {
                            #if DEBUG
                            Debug.WriteLine($"[PS OUTPUT] {info}");
                            #endif
                        }
                    }

                    // Capture any regular output
                    if (results != null && results.Count > 0)
                    {
                        foreach (var output in results)
                        {
                            #if DEBUG
                            Debug.WriteLine($"[PS RESULT] {output}");
                            #endif
                        }
                    }

                    if (ps.HadErrors)
                    {
                        #if DEBUG
                        Debug.WriteLine($"[LEQ TOGGLE] Errors:");
                        #endif
                        foreach (var error in ps.Streams.Error)
                        {
                            #if DEBUG
                            Debug.WriteLine($"LEQ TOGGLE ERROR DETAIL: {error.ToString()}");
                            Debug.WriteLine($"LEQ TOGGLE ERROR EXCEPTION: {error.Exception?.Message}");
                            if (error.Exception != null)
                            {
                                Debug.WriteLine($"LEQ TOGGLE ERROR STACK TRACE: {error.Exception.StackTrace}");
                            }
                            #endif
                        }
                        return (bool?)null;
                    }

                    #if DEBUG
                    Debug.WriteLine($"\u2713 LEQ {(newEnabled ? "enabled" : "disabled")}");
                    #endif
                    
#if DEBUG
                    // Dump full registry state for debugging
                    var dumpScript = @"
                        param($deviceGuid)
                        $fxPath = ""HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\$deviceGuid\FxProperties""

                        Write-Output ""`n=== FULL REGISTRY DUMP for $deviceGuid ===""

                        if (-not (Test-Path -LiteralPath $fxPath)) {
                            Write-Output ""FxProperties path does not exist!""
                            return
                        }

                        $props = Get-ItemProperty -LiteralPath $fxPath
                        $allProps = $props.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' }

                        Write-Output ""Total properties: $($allProps.Count)""
                        Write-Output """"

                        foreach ($prop in $allProps | Sort-Object Name) {
                            $val = $prop.Value
                            if ($val -is [byte[]]) {
                                $hexStr = ($val | ForEach-Object { '{0:x2}' -f $_ }) -join ','
                                Write-Output ""$($prop.Name) = [BINARY] $hexStr""
                            } else {
                                Write-Output ""$($prop.Name) = $val""
                            }
                        }

                        Write-Output ""`n=== END REGISTRY DUMP ===""
                    ";

                    ps.Commands.Clear();
                    ps.AddScript(dumpScript).AddParameter("deviceGuid", deviceId);
                    var dumpResults = ps.Invoke();
                    foreach (var line in dumpResults)
                    {
                        Debug.WriteLine($"[REG DUMP] {line}");
                    }
#endif

                    return (bool?)newEnabled;
                }
            }
            catch (Exception)
            {
                #if DEBUG
                Debug.WriteLine("[LEQ TOGGLE] Exception occurred");
                #endif
                return (bool?)null;
            }
        });
    }


    public async Task<bool?> GetLeqStateAsync(string deviceId)
    {
        if (string.IsNullOrWhiteSpace(deviceId)) return null;
        return await Task.Run(() => ReadLeqStateFromRegistry(deviceId));
    }

    public async Task<bool?> GetEapoStatusAsync(string deviceId)
    {
        #if DEBUG
        Debug.WriteLine($"AudioService: reading E-APO status via embedded scripts for device {deviceId}");
        #endif

        return await Task.Run(() =>
        {
            try
            {
                using (var psWrapper = CreateConfiguredPowerShell())
                {
                    var ps = psWrapper.PowerShell;
                    
                    ps.AddCommand("Get-EapoStatus")
                      .AddParameter("DeviceGuid", deviceId);

                    var results = InvokeWithTimeout(ps);

                    // Capture Write-Host output from Information stream
                    if (ps.Streams.Information.Count > 0)
                    {
                        foreach (var info in ps.Streams.Information)
                        {
                            Debug.WriteLine($"[EAPO PS]: {info}");
                        }
                    }

                    // Check for PowerShell errors
                    if (ps.Streams.Error.Count > 0)
                    {
                        Debug.WriteLine($"[EAPO STATUS] PowerShell error count: {ps.Streams.Error.Count}");

                        foreach (var error in ps.Streams.Error)
                        {
                            Debug.WriteLine($"EAPO STATUS ERROR DETAIL: {error.ToString()}");
                            Debug.WriteLine($"EAPO STATUS ERROR EXCEPTION: {error.Exception?.Message}");
                            if (error.Exception != null)
                            {
                                Debug.WriteLine($"EAPO STATUS ERROR STACK TRACE: {error.Exception.StackTrace}");
                            }
                        }

                        var allErrors = string.Join("\n", ps.Streams.Error.Select(e => e.ToString()));
                        throw new Exception($"PowerShell execution failed in Get-EapoStatus:\n{allErrors}");
                    }

                    // If we get here, error count was 0, process results
                    if (results != null && results.Count > 0)
                    {
                        var firstResult = results[0];
                        if (firstResult?.BaseObject is bool result)
                        {
                            return result;
                        }
                    }

                    return (bool?)null;
                }
            }
            catch (Exception)
            {
                #if DEBUG
                Debug.WriteLine("AudioService: GetEapoStatus unexpected error");
                #endif
                return null;
            }
        });
    }


    public async Task<bool> InstallLeqAsync(string deviceId, bool cleanInstall = false)
    {
        return await Task.Run(() =>
        {
            const int initialReleaseTime = 4;

            try
            {
                using (var psWrapper = CreateConfiguredPowerShell())
                {
                    var ps = psWrapper.PowerShell;

                    var targetDevice = FindDeviceByGuid(ps, deviceId);
                    if (targetDevice == null)
                    {
                        #if DEBUG
                        Debug.WriteLine($"[LEQ INSTALL] Device not found: {deviceId}");
                        #endif
                        return false;
                    }

                    // Install LEQ
                    ps.AddCommand("Install-LEQRegistry")
                      .AddParameter("Device", targetDevice)
                      .AddParameter("ReleaseTime", initialReleaseTime);

                    if (cleanInstall)
                    {
                        ps.AddParameter("Force", true);
                    }

                    ps.Invoke();

                    if (ps.HadErrors)
                    {
                        #if DEBUG
                        Debug.WriteLine($"[LEQ INSTALL] Errors:");
                        #endif
                        foreach (var error in ps.Streams.Error)
                        {
                            #if DEBUG
                            Debug.WriteLine($"LEQ INSTALL ERROR DETAIL: {error.ToString()}");
                            Debug.WriteLine($"LEQ INSTALL ERROR EXCEPTION: {error.Exception?.Message}");
                            if (error.Exception != null)
                            {
                                Debug.WriteLine($"LEQ INSTALL ERROR STACK TRACE: {error.Exception.StackTrace}");
                            }
                            #endif
                        }
                        return false;
                    }

                    #if DEBUG
                    Debug.WriteLine($"[LEQ INSTALL] Success");
                    #endif
                    return true;
                }
            }
            catch (Exception)
            {
                #if DEBUG
                Debug.WriteLine("[LEQ INSTALL] Exception occurred");
                #endif
                return false;
            }
        });
    }

    private static readonly (string guid, string name)[] RequiredClsids = new[]
    {
        ("{62dc1a93-ae24-464c-a43e-452f824c4250}", "LEQ APO"),
        ("{637c490d-eee3-4c0a-973f-371958802da2}", "Enhancement APO"),
        ("{13AB3EBD-137E-4903-9D89-60BE8277FD17}", "WM GFX APO"),
        ("{C9453E73-8C5C-4463-9984-AF8BAB2F5447}", "WM LFX APO"),
    };

    private const string EnhancementDllPath = @"C:\WINDOWS\System32\WMALFXGFXDSP.dll";

    /// <summary>
    /// Quick read-only check if any CLSIDs are missing or misconfigured.
    /// Does NOT fix anything — just returns the names of broken CLSIDs.
    /// </summary>
    public string[] CheckClsidHealth()
    {
        var broken = new List<string>();
        foreach (var (guid, name) in RequiredClsids)
        {
            var keyPath = $@"SOFTWARE\Classes\CLSID\{guid}\InprocServer32";
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(keyPath);
                if (key == null)
                {
                    broken.Add(name);
                    continue;
                }
                var defaultVal = key.GetValue("")?.ToString();
                var threading = key.GetValue("ThreadingModel")?.ToString();
                if (!string.Equals(defaultVal, EnhancementDllPath, StringComparison.OrdinalIgnoreCase) ||
                    threading != "Both")
                {
                    broken.Add(name);
                }
            }
            catch
            {
                broken.Add(name);
            }
        }
        return broken.ToArray();
    }

    public async Task<(string[] missing, string[] fixed_, string[] failed)> RepairClsidsAsync()
    {
        return await Task.Run(() =>
        {
            try
            {
                using (var psWrapper = CreateConfiguredPowerShell())
                {
                    var ps = psWrapper.PowerShell;
                    ps.AddCommand("Repair-LEQClsids");

                    var results = ps.Invoke();

                    if (results?.Count > 0 &&
                        results[0]?.BaseObject is System.Collections.Hashtable ht)
                    {
                        return (
                            ParseStringArray(ht["Missing"]),
                            ParseStringArray(ht["Fixed"]),
                            ParseStringArray(ht["Failed"])
                        );
                    }

                    return (Array.Empty<string>(), Array.Empty<string>(), Array.Empty<string>());
                }
            }
            catch (Exception)
            {
                throw;
            }
        });
    }

    public async Task<bool> FixDeviceCompositeFxAsync(string deviceGuid)
    {
        return await Task.Run(() =>
        {
            try
            {
                using (var psWrapper = CreateConfiguredPowerShell())
                {
                    var ps = psWrapper.PowerShell;
                    ps.AddCommand("Clear-CompositeFxKeys")
                      .AddParameter("DeviceGuid", deviceGuid);

                    var results = ps.Invoke();

                    if (ps.HadErrors)
                    {
#if DEBUG
                        foreach (var error in ps.Streams.Error)
                            Debug.WriteLine($"[FIX COMPOSITEFX] Error: {error}");
#endif
                        return false;
                    }

                    // Clear-CompositeFxKeys returns $true on success
                    if (results?.Count > 0)
                    {
                        var val = results[0]?.BaseObject;
                        if (val is bool b) return b;
                    }

                    return true;
                }
            }
            catch
            {
                return false;
            }
        });
    }

    private static string[] ParseStringArray(object? value)
    {
        if (value is object[] arr)
            return arr.Select(o => o?.ToString() ?? "").Where(s => !string.IsNullOrEmpty(s)).ToArray();
        if (value is string s && !string.IsNullOrEmpty(s))
            return new[] { s };
        return Array.Empty<string>();
    }

    public async Task<(bool success, string? reason)> RestartAudioServiceAsync()
    {
        var script = @"
            $result = @{ Success = $false; Reason = '' }
            try {
                Restart-Service audiosrv -Force -ErrorAction Stop
                $result.Success = $true
            } catch {
                $result.Reason = ""Failed to restart audio service: $($_.Exception.Message)""
            }
            return $result
        ";

        var sessionState = InitialSessionState.CreateDefault();
        sessionState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Unrestricted;

        var runspace = RunspaceFactory.CreateRunspace(sessionState);
        runspace.Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(script);

        try
        {
            var results = await Task.Run(() => shell.Invoke());

            if (shell.Streams.Error?.Count > 0)
            {
                var errorMessage = string.Join(" | ", shell.Streams.Error.Select(err => err.ToString()));
                #if DEBUG
                Debug.WriteLine($"AudioService: RestartAudioService reported errors: {errorMessage}");
                #endif
                return (false, errorMessage);
            }

            if (results is null || results.Count == 0)
            {
                return (false, "Audio restart script returned no result.");
            }

            var hashtable = results[0]?.BaseObject as System.Collections.Hashtable;
            if (hashtable is null)
            {
                return (false, "Audio restart script returned an unexpected result.");
            }

            var success = hashtable["Success"] is bool s && s;
            var reason = hashtable["Reason"]?.ToString();
            return (success, string.IsNullOrWhiteSpace(reason) ? null : reason);
        }
        finally
        {
            runspace.Dispose();
        }
    }

    public async Task<(bool success, string? message)> ResetDeviceAsync(string deviceGuid, string deviceName)
    {
        var script = @"
            param($deviceGuid, $deviceName)
            try {
                # Step 1: Read adapter device instance ID from MMDevices registry
                # Property {b3f8fa53-0004-438e-9003-51a46e139bfc},2 contains the parent adapter
                # instance path with a {N}. prefix, e.g. ""{1}.USB\VID_041E&PID_329B&MI_00\9&71A7601&A&0000""
                $propsPath = ""HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\$deviceGuid\Properties""

                if (-not (Test-Path -LiteralPath $propsPath)) {
                    return @{ Success = $false; Message = 'Device properties not found in registry.' }
                }

                $props = Get-ItemProperty -LiteralPath $propsPath -ErrorAction SilentlyContinue
                $adapterPropKey = '{b3f8fa53-0004-438e-9003-51a46e139bfc},2'

                $instanceId = $null

                if ($props -and $props.PSObject.Properties.Name -contains $adapterPropKey) {
                    $adapterPath = $props.$adapterPropKey
                    if ($adapterPath) {
                        # Strip the {N}. prefix to get the raw PnP instance ID
                        # ""{1}.USB\VID_041E&PID_329B..."" -> ""USB\VID_041E&PID_329B...""
                        $instanceId = $adapterPath -replace '^\{[0-9]+\}\.', ''
                    }
                }

                if (-not $instanceId) {
                    return @{ Success = $false; Message = 'Could not determine hardware device instance ID for this audio endpoint.' }
                }

                # Find a valid PnP device to remove using multi-tier lookup
                # (composite USB interface children are often invisible to Get-PnpDevice)
                $targetId = $null

                # Tier 1: Get-PnpDevice exact match (fast, works for most devices)
                $pnpDevice = Get-PnpDevice -InstanceId $instanceId -ErrorAction SilentlyContinue
                if ($pnpDevice) {
                    $targetId = $instanceId
                }

                # Tier 2: pnputil /enum-devices has broader visibility (hidden/software devices)
                if (-not $targetId) {
                    $enumResult = & pnputil /enum-devices /instanceid ""$instanceId"" 2>&1
                    $found = $enumResult | Select-String -Pattern ([regex]::Escape($instanceId)) -Quiet
                    if ($found) {
                        $targetId = $instanceId
                    }
                }

                # Tier 3: For composite USB devices (MI_xx), try the parent composite device
                if (-not $targetId -and $instanceId -match '&MI_[0-9A-Fa-f]{2}') {
                    $parentId = $instanceId -replace '&MI_[0-9A-Fa-f]{2}', ''

                    $parentDevice = Get-PnpDevice -InstanceId $parentId -ErrorAction SilentlyContinue
                    if ($parentDevice) {
                        $targetId = $parentId
                    } else {
                        $enumParent = & pnputil /enum-devices /instanceid ""$parentId"" 2>&1
                        $foundParent = $enumParent | Select-String -Pattern ([regex]::Escape($parentId)) -Quiet
                        if ($foundParent) {
                            $targetId = $parentId
                        }
                    }
                }

                # Tier 4: Try pnputil directly — it will report its own error if device not found
                if (-not $targetId) {
                    $targetId = $instanceId
                }

                # Step 1.5: Delete friendly name so Windows regenerates the default on re-detection
                $friendlyNameKey = '{a45c254e-df1c-4efd-8020-67d146a850e0},2'
                if ($props -and $props.PSObject.Properties.Name -contains $friendlyNameKey) {
                    try {
                        Remove-ItemProperty -LiteralPath $propsPath -Name $friendlyNameKey -ErrorAction Stop
                    } catch {
                        # Non-fatal: device still gets reset even if rename fails
                    }
                }

                # Step 2: Remove device and re-scan (more reliable than restart)
                $removeResult = & pnputil /remove-device ""$targetId"" 2>&1
                $removeExitCode = $LASTEXITCODE
                $removeOutput = ($removeResult | Out-String).Trim()

                if ($removeExitCode -eq 0 -or $removeOutput -match 'success') {
                    Start-Sleep -Seconds 2
                    & pnputil /scan-devices 2>&1 | Out-Null
                    Start-Sleep -Seconds 1
                } else {
                    # Remove failed — fall back to restart + scan
                    $restartResult = & pnputil /restart-device ""$targetId"" 2>&1
                    $restartExitCode = $LASTEXITCODE
                    $restartOutput = ($restartResult | Out-String).Trim()

                    if ($restartExitCode -ne 0 -and $restartOutput -notmatch 'success') {
                        return @{ Success = $false; Message = ""Device reset failed. Remove: $removeOutput. Restart: $restartOutput"" }
                    }

                    Start-Sleep -Seconds 2
                    & pnputil /scan-devices 2>&1 | Out-Null
                    Start-Sleep -Seconds 1
                }

                return @{ Success = $true; Message = $null }
            } catch {
                return @{ Success = $false; Message = $_.Exception.Message }
            }
        ";

        var sessionState = InitialSessionState.CreateDefault();
        sessionState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Unrestricted;

        var runspace = RunspaceFactory.CreateRunspace(sessionState);
        runspace.Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(script)
             .AddParameter("deviceGuid", deviceGuid)
             .AddParameter("deviceName", deviceName);

        try
        {
            var results = await Task.Run(() => shell.Invoke());

            if (shell.Streams.Error?.Count > 0)
            {
                var errorMessage = string.Join(" | ", shell.Streams.Error.Select(err => err.ToString()));
                #if DEBUG
                Debug.WriteLine($"AudioService: ResetDeviceAsync reported errors: {errorMessage}");
                #endif
                return (false, errorMessage);
            }

            if (results is null || results.Count == 0)
            {
                return (false, "No result returned from device reset script.");
            }

            var result = results[0];
            if (result?.BaseObject is System.Collections.Hashtable ht)
            {
                var success = ht["Success"] is bool s && s;
                var message = ht["Message"]?.ToString();
                return (success, message);
            }

            return (false, "Unexpected result format from device reset script.");
        }
        finally
        {
            runspace.Dispose();
        }
    }

    public async Task<List<(string Guid, string Name)>> GetSiblingEndpointsAsync(string deviceGuid)
    {
        var script = @"
            param($deviceGuid)
            $renderRoot = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render'
            $propsPath = ""$renderRoot\$deviceGuid\Properties""
            $adapterPropKey = '{b3f8fa53-0004-438e-9003-51a46e139bfc},2'

            $props = Get-ItemProperty -LiteralPath $propsPath -ErrorAction SilentlyContinue
            if (-not $props -or -not ($props.PSObject.Properties.Name -contains $adapterPropKey)) {
                return @()
            }

            $instanceId = $props.$adapterPropKey -replace '^\{[0-9]+\}\.', ''
            # Strip MI_xx to get the base composite device ID
            $baseId = $instanceId -replace '&MI_[0-9A-Fa-f]{2}', ''

            $siblings = @()
            $deviceKeys = Get-ChildItem -LiteralPath $renderRoot -ErrorAction SilentlyContinue
            foreach ($dk in $deviceKeys) {
                $otherGuid = Split-Path -Leaf $dk.PSPath
                if ($otherGuid -eq $deviceGuid) { continue }

                # Only include active endpoints (DeviceState = 1)
                $otherDeviceInfo = Get-ItemProperty -LiteralPath $dk.PSPath -ErrorAction SilentlyContinue
                $otherState = if ($otherDeviceInfo -and $otherDeviceInfo.PSObject.Properties.Name -contains 'DeviceState') { $otherDeviceInfo.DeviceState } else { 0 }
                if ($otherState -ne 1) { continue }

                $otherPropsPath = Join-Path $dk.PSPath 'Properties'
                $otherProps = Get-ItemProperty -LiteralPath $otherPropsPath -ErrorAction SilentlyContinue
                if (-not $otherProps -or -not ($otherProps.PSObject.Properties.Name -contains $adapterPropKey)) { continue }

                $otherInstanceId = $otherProps.$adapterPropKey -replace '^\{[0-9]+\}\.', ''
                $otherBaseId = $otherInstanceId -replace '&MI_[0-9A-Fa-f]{2}', ''

                if ($otherBaseId -eq $baseId) {
                    $friendlyNameKey = '{a45c254e-df1c-4efd-8020-67d146a850e0},2'
                    $otherName = if ($otherProps.PSObject.Properties.Name -contains $friendlyNameKey) {
                        $otherProps.$friendlyNameKey
                    } else { 'Unknown' }

                    $siblings += [PSCustomObject]@{ Guid = $otherGuid; Name = [string]$otherName }
                }
            }
            return $siblings
        ";

        var sessionState = InitialSessionState.CreateDefault();
        sessionState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Unrestricted;

        var runspace = RunspaceFactory.CreateRunspace(sessionState);
        runspace.Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(script)
             .AddParameter("deviceGuid", deviceGuid);

        try
        {
            var results = await Task.Run(() => shell.Invoke());
            var siblings = new List<(string Guid, string Name)>();

            if (results != null)
            {
                foreach (var r in results)
                {
                    if (r?.Properties["Guid"] != null)
                    {
                        var guid = r.Properties["Guid"]?.Value?.ToString() ?? "";
                        var name = r.Properties["Name"]?.Value?.ToString() ?? "Unknown";
                        if (!string.IsNullOrEmpty(guid))
                            siblings.Add((guid, name));
                    }
                }
            }

            return siblings;
        }
        finally
        {
            runspace.Dispose();
        }
    }

    public async Task SetReleaseTimeAsync(string deviceId, double value)
    {
        await Task.Run(() =>
        {
            int rtValue = Math.Clamp((int)Math.Round(value), 2, 7);

            try
            {
                using (var psWrapper = CreateConfiguredPowerShell())
                {
                    var ps = psWrapper.PowerShell;
                    
                    var script = @"
                        param($guid, $rt)
                        Set-ReleaseTime -DeviceGuid $guid -ReleaseTime $rt
                    ";

                    ps.AddScript(script)
                      .AddParameter("guid", deviceId)
                      .AddParameter("rt", rtValue);

                    var results = InvokeWithTimeout(ps);

                    foreach (var output in results)
                    {
                        #if DEBUG
                        Debug.WriteLine($"[RT] {output}");
                        #endif
                    }

                    if (ps.HadErrors)
                    {
                        foreach (var error in ps.Streams.Error)
                        {
                            #if DEBUG
                            Debug.WriteLine($"RT ERROR DETAIL: {error.ToString()}");
                            Debug.WriteLine($"RT ERROR EXCEPTION: {error.Exception?.Message}");
                            if (error.Exception != null)
                            {
                                Debug.WriteLine($"RT ERROR STACK TRACE: {error.Exception.StackTrace}");
                            }
                            #endif
                        }
                    }
                }
            }
            catch (Exception)
            {
                #if DEBUG
                Debug.WriteLine("[RT] Exception occurred");
                #endif
            }
        });
    }

    private static bool ParseBoolean(object? value)
    {
        if (value is bool boolValue)
        {
            return boolValue;
        }

        if (bool.TryParse(value?.ToString(), out var parsed))
        {
            return parsed;
        }

        return false;
    }

    /// <summary>
    /// Reads the LEQ enabled/disabled state directly from the registry (no PowerShell).
    /// Returns true if enabled, false if disabled, null if the key doesn't exist or can't be read.
    /// </summary>
    public static bool? ReadLeqStateFromRegistry(string deviceGuid)
    {
        const string renderRoot = @"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render";
        const string leqKey = "{fc52a749-4be9-4510-896e-966ba6525980},3";

        try
        {
            using var fxKey = Registry.LocalMachine.OpenSubKey(
                $@"{renderRoot}\{deviceGuid}\FxProperties", false);
            if (fxKey == null) return null;

            var value = fxKey.GetValue(leqKey) as byte[];
            if (value == null || value.Length < 10) return null;

            // Bytes 8-9 == 0xFF,0xFF means enabled (matches LEQ-Engine.ps1 line 564)
            return value[8] == 0xFF && value[9] == 0xFF;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Reads the DeviceState for a render endpoint directly from the registry (no PowerShell).
    /// Returns 1 for Active/Ready, 2 for Disabled, 4 for NotPresent, 8 for Unplugged.
    /// Returns 0 if the key doesn't exist or can't be read (0 is never a valid DeviceState).
    /// </summary>
    public static int ReadDeviceStateFromRegistry(string deviceGuid)
    {
        const string renderRoot = @"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render";

        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(
                $@"{renderRoot}\{deviceGuid}", false);
            if (key == null) return 0;

            var val = key.GetValue("DeviceState");
            if (val is int intVal) return intVal;
            return 0;
        }
        catch
        {
            return 0;
        }
    }

    private static string LoadEmbeddedScript(string resourceName)
    {
        var assembly = typeof(AudioService).Assembly;
        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null)
        {
            var availableResources = assembly.GetManifestResourceNames();
            #if DEBUG
            Debug.WriteLine($"Embedded resource '{resourceName}' not found.");
            Debug.WriteLine($"Available resources: {string.Join(", ", availableResources)}");
            #endif
            throw new InvalidOperationException($"Embedded resource '{resourceName}' not found. Available: {string.Join(", ", availableResources)}");
        }

        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
