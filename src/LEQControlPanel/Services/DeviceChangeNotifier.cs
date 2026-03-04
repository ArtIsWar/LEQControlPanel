// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System;
using System.Diagnostics;
using System.Threading;

namespace LEQControlPanel.Services;

internal enum DeviceChangeType { DeviceAdded, DeviceRemoved, StateChanged, DefaultDeviceChanged }

internal class DeviceChangedEventArgs : EventArgs
{
    public DeviceChangeType ChangeType { get; }
    public string? DeviceId { get; }
    /// <summary>Device state for StateChanged events. 1=Active, 2=Disabled, 4=NotPresent, 8=Unplugged.</summary>
    public int? NewState { get; }

    public DeviceChangedEventArgs(DeviceChangeType changeType, string? deviceId = null, int? newState = null)
    {
        ChangeType = changeType;
        DeviceId = deviceId;
        NewState = newState;
    }
}

/// <summary>
/// Listens for Windows audio device changes using registry monitoring via
/// <see cref="RegistryDeviceMonitor"/> (RegNotifyChangeKeyValue). Pure managed code
/// with no COM dependency — immune to third-party driver crashes that corrupt shared
/// audio state and deliver uncatchable AccessViolationException via COM RPC proxy.
/// </summary>
internal sealed class DeviceChangeNotifier : IDisposable
{
    private RegistryDeviceMonitor? _registryMonitor;
    private volatile bool _disposed;
    private int _suspendCount;
    private long _resumeGraceUntilTicks; // Stopwatch ticks — events suppressed until this time
    private long _lastNotificationTicks; // For debounce (collapse burst registry notifications)

    /// <summary>Grace period after Resume() during which events are silently dropped.
    /// Prevents the cascade of add/remove/state-change events that Windows fires
    /// as devices re-enumerate after an audio service restart.</summary>
    private static readonly long GracePeriodTicks = Stopwatch.Frequency * 3; // 3 seconds

    /// <summary>Debounce window for registry notifications. bWatchSubtree = true on the
    /// Render hive fires on every property change for every device — this collapses
    /// burst notifications into a single event.</summary>
    private static readonly long DebounceWindowTicks = Stopwatch.Frequency / 4; // 250ms

    public event EventHandler<DeviceChangedEventArgs>? DeviceChanged;

    public DeviceChangeNotifier()
    {
        try
        {
            _registryMonitor = new RegistryDeviceMonitor();
            _registryMonitor.Changed += OnRegistryChanged;
            _registryMonitor.Start();
            Debug.WriteLine("[DeviceChangeNotifier] Registry monitor started");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[DeviceChangeNotifier] Registry monitor failed to start: {ex.Message}");
        }
    }

    private void OnRegistryChanged()
    {
        if (_disposed || _suspendCount > 0) return;

        // Suppress events during the post-resume grace period
        if (Stopwatch.GetTimestamp() < Interlocked.Read(ref _resumeGraceUntilTicks)) return;

        // Debounce: collapse burst notifications from the registry watcher
        var now = Stopwatch.GetTimestamp();
        var last = Interlocked.Read(ref _lastNotificationTicks);
        if (now - last < DebounceWindowTicks) return;

        Interlocked.Exchange(ref _lastNotificationTicks, now);

        try
        {
            DeviceChanged?.Invoke(this, new DeviceChangedEventArgs(DeviceChangeType.StateChanged));
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[DeviceChangeNotifier] Subscriber error: {ex.Message}");
        }
    }

    /// <summary>
    /// Stops the registry monitor before an audio service restart or device operation.
    /// </summary>
    /// <remarks>
    /// Re-entrant — multiple callers can nest suspend/resume pairs. Notifications
    /// are only stopped on the first suspend (0→1) and only restarted when the last
    /// resume brings the count back to 0.
    /// </remarks>
    public void Suspend()
    {
        if (_disposed) return;

        var count = Interlocked.Increment(ref _suspendCount);
        if (count > 1)
        {
            Debug.WriteLine($"[DeviceChangeNotifier] Suspend nested (count={count})");
            return;
        }

        _registryMonitor?.Stop();
        Debug.WriteLine($"[DeviceChangeNotifier] Suspended (count={count})");
    }

    /// <summary>
    /// Restarts the registry monitor after an audio service restart or device operation.
    /// </summary>
    /// <remarks>
    /// Re-entrant — only restarts when the suspend count reaches 0.
    /// </remarks>
    public void Resume()
    {
        if (_disposed) return;

        var count = Interlocked.Decrement(ref _suspendCount);
        if (count < 0)
        {
            Interlocked.Exchange(ref _suspendCount, 0);
            Debug.WriteLine("[DeviceChangeNotifier] Resume called more times than Suspend — clamped to 0");
            return;
        }
        if (count > 0)
        {
            Debug.WriteLine($"[DeviceChangeNotifier] Resume nested (count={count})");
            return;
        }

        // count == 0: last resume — restart notification source

        // Start grace period — suppress the cascade of events as devices re-enumerate
        Interlocked.Exchange(ref _resumeGraceUntilTicks, Stopwatch.GetTimestamp() + GracePeriodTicks);

        try
        {
            _registryMonitor?.Start();
            Debug.WriteLine("[DeviceChangeNotifier] Registry monitor resumed");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[DeviceChangeNotifier] Registry monitor resume failed: {ex.Message}");
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_registryMonitor != null)
        {
            _registryMonitor.Changed -= OnRegistryChanged;
            _registryMonitor.Dispose();
            _registryMonitor = null;
        }

        GC.SuppressFinalize(this);
    }
}
