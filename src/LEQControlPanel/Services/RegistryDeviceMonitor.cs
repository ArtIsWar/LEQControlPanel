// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace LEQControlPanel.Services;

/// <summary>
/// Watches the MMDevices\Audio\Render registry hive for changes using
/// RegNotifyChangeKeyValue. This provides push-based device change detection
/// without relying on COM callbacks, which can crash if third-party audio
/// drivers corrupt shared state (e.g., Elgato's WindowsAudioRouterApi.dll).
/// </summary>
internal sealed class RegistryDeviceMonitor : IDisposable
{
    // --- P/Invoke ---

    [DllImport("advapi32.dll", EntryPoint = "RegOpenKeyExW", CharSet = CharSet.Unicode)]
    private static extern int RegOpenKeyEx(IntPtr hKey, string subKey, uint options, uint samDesired, out IntPtr phkResult);

    [DllImport("advapi32.dll")]
    private static extern int RegCloseKey(IntPtr hKey);

    [DllImport("advapi32.dll")]
    private static extern int RegNotifyChangeKeyValue(IntPtr hKey, bool bWatchSubtree, uint dwNotifyFilter, IntPtr hEvent, bool fAsynchronous);

    [DllImport("kernel32.dll", EntryPoint = "CreateEventW", CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateEvent(IntPtr lpEventAttributes, bool bManualReset, bool bInitialState, string? lpName);

    [DllImport("kernel32.dll")]
    private static extern uint WaitForMultipleObjects(uint nCount, IntPtr[] lpHandles, bool bWaitAll, uint dwMilliseconds);

    [DllImport("kernel32.dll")]
    private static extern bool CloseHandle(IntPtr hObject);

    // --- Constants ---

    private static readonly IntPtr HKEY_LOCAL_MACHINE = new(unchecked((int)0x80000002));
    private const uint KEY_NOTIFY = 0x0010;
    private const uint KEY_READ = 0x20019;
    private const uint REG_NOTIFY_CHANGE_NAME = 0x1;      // Subkey add/delete (device add/remove)
    private const uint REG_NOTIFY_CHANGE_LAST_SET = 0x4;   // Value changes (state, properties)
    private const uint WAIT_OBJECT_0 = 0;
    private const uint WAIT_FAILED = 0xFFFFFFFF;
    private const uint INFINITE = 0xFFFFFFFF;

    private const string RenderRoot = @"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render";

    // --- State ---

    private Thread? _watchThread;
    private readonly ManualResetEventSlim _stopEvent = new(false);
    private volatile bool _disposed;

    /// <summary>Raised on a background thread when any audio render device registry state changes.</summary>
    public event Action? Changed;

    public void Start()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(RegistryDeviceMonitor));
        if (_watchThread != null) return; // already running

        _stopEvent.Reset();
        _watchThread = new Thread(WatchLoop)
        {
            Name = "RegistryDeviceMonitor",
            IsBackground = true
        };
        _watchThread.Start();
    }

    public void Stop()
    {
        if (_watchThread == null) return;

        _stopEvent.Set();
        _watchThread.Join(timeout: TimeSpan.FromSeconds(5));
        _watchThread = null;
    }

    private void WatchLoop()
    {
        // Create a Win32 manual-reset event for stop signaling
        IntPtr stopHandle = _stopEvent.WaitHandle.SafeWaitHandle.DangerousGetHandle();

        while (!_disposed && !_stopEvent.IsSet)
        {
            IntPtr hKey = IntPtr.Zero;
            IntPtr hEvent = IntPtr.Zero;

            try
            {
                // Open the registry key
                int rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, RenderRoot, 0, KEY_NOTIFY | KEY_READ, out hKey);
                if (rc != 0)
                {
                    Debug.WriteLine($"[RegistryDeviceMonitor] RegOpenKeyEx failed: 0x{rc:X}");
                    // Wait before retrying to avoid tight loop
                    if (_stopEvent.Wait(TimeSpan.FromSeconds(5))) return;
                    continue;
                }

                // Create a notification event (auto-reset)
                hEvent = CreateEvent(IntPtr.Zero, bManualReset: false, bInitialState: false, null);
                if (hEvent == IntPtr.Zero)
                {
                    Debug.WriteLine("[RegistryDeviceMonitor] CreateEvent failed");
                    if (_stopEvent.Wait(TimeSpan.FromSeconds(5))) return;
                    continue;
                }

                // Register for change notification (one-shot, async)
                rc = RegNotifyChangeKeyValue(hKey, bWatchSubtree: true,
                    REG_NOTIFY_CHANGE_NAME | REG_NOTIFY_CHANGE_LAST_SET,
                    hEvent, fAsynchronous: true);
                if (rc != 0)
                {
                    Debug.WriteLine($"[RegistryDeviceMonitor] RegNotifyChangeKeyValue failed: 0x{rc:X}");
                    if (_stopEvent.Wait(TimeSpan.FromSeconds(5))) return;
                    continue;
                }

                // Wait for either the registry change event or the stop signal
                var handles = new[] { hEvent, stopHandle };
                uint result = WaitForMultipleObjects(2, handles, bWaitAll: false, INFINITE);

                if (result == WAIT_OBJECT_0)
                {
                    // Registry changed — fire event
                    if (!_disposed && !_stopEvent.IsSet)
                    {
                        try
                        {
                            Changed?.Invoke();
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[RegistryDeviceMonitor] Subscriber error: {ex.Message}");
                        }
                    }
                    // Loop to re-register (notification is one-shot)
                }
                else if (result == WAIT_OBJECT_0 + 1)
                {
                    // Stop requested
                    return;
                }
                else
                {
                    Debug.WriteLine($"[RegistryDeviceMonitor] WaitForMultipleObjects returned 0x{result:X}");
                    if (_stopEvent.Wait(TimeSpan.FromSeconds(1))) return;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[RegistryDeviceMonitor] Watch loop error: {ex.Message}");
                if (_stopEvent.Wait(TimeSpan.FromSeconds(5))) return;
            }
            finally
            {
                if (hEvent != IntPtr.Zero) CloseHandle(hEvent);
                if (hKey != IntPtr.Zero) RegCloseKey(hKey);
            }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        Stop();
        _stopEvent.Dispose();
    }
}
