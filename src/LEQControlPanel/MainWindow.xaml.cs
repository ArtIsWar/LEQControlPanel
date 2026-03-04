// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using Microsoft.Win32;
using Application = System.Windows.Application;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using LEQControlPanel.Dialogs;
using LEQControlPanel.Services;
using LEQControlPanel.Models;
using LEQControlPanel.Utilities;
using LEQControlPanel.Windows;
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using Color = System.Windows.Media.Color;

namespace LEQControlPanel
{
    public partial class MainWindow : Window
    {

        private NotifyIcon _trayIcon = null!;
        private AudioService _audioService = null!;
        private readonly DispatcherTimer _debounceTimer = null!;
        private DeviceChangeNotifier? _deviceChangeNotifier;
        private readonly DispatcherTimer _deviceChangeDebounceTimer;
        private readonly List<DeviceChangedEventArgs> _pendingDeviceChanges = new();
        private bool _isRefreshing;
        private volatile bool _isManuallyResetting;
        private bool _clsidsBroken;
        internal List<AudioDevice> _devices = new();
        private bool _isSliderUpdating;
        private volatile bool _isInitialized;
        private volatile bool _isExiting;

        private int _lastLoggedReleaseValue = -1;
        private ContextMenuStrip _trayContextMenu = null!;
        private ToolStripMenuItem _runAtStartupItem = null!;
        private ToolStripMenuItem _toggleLeqMenuItem = null!;
        private ToolStripMenuItem _deviceFriendlyNameMenuItem = null!;
        private ToolStripMenuItem _deviceDescriptorMenuItem = null!;
        private ToolStripMenuItem _restartAudioServiceItem = null!;
        private ToolStripMenuItem _alwaysOnTopTrayItem = null!;
        private ToolStripMenuItem _desktopShortcutTrayItem = null!;
        private const string RunRegistryPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunValueName = "LEQControlPanel";
        private const string SettingsRegistryPath = @"Software\LEQControlPanel";
        private bool _runAtStartupEnabled;
        internal bool _startMinimized = false;
        private CloseBehavior _closeBehavior = CloseBehavior.ExitApplication;
        private bool _closeBehaviorRemembered;
        private bool _skipEapoWarning;
        private bool _skipResetDeviceWarning;
        private bool _alwaysOnTop = false;

        private bool _isUpdatingIndicator;
        private bool _isUpdatingEapo;
        private bool? _lastLeqState;
        private bool _isAdmin;
        private bool _isEapoInstalled;
        private string _lastSelectedDeviceGuid = "";

        private bool _atkInstalled;
        private string? _atkExePath;
        private bool _glowAnimationsSuppressed;
        internal SplashWindow? SplashScreen { get; set; }

        public MainWindow()
        {
            _debounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
            _debounceTimer.Tick += DebounceTimer_Tick;

            _deviceChangeDebounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(2000) };
            _deviceChangeDebounceTimer.Tick += DeviceChangeDebounceTimer_Tick;

            InitializeComponent();
            LoadApplicationSettings();

            // Handle start minimized
            if (_startMinimized)
            {
                WindowState = WindowState.Minimized;
                Hide();
            }

            // Check admin status (fast, no UI)
            using (var identity = WindowsIdentity.GetCurrent())
                _isAdmin = new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Finish initialization that was deferred from constructor to keep splash responsive
                if (!_isAdmin)
                {
                    StyledMessageBox.ShowWarning(
                        "LEQ Control Panel requires Administrator privileges to manage audio settings.\n\n" +
                        "LEQ toggle and Release Time controls will not function without admin rights.\n\n" +
                        "Please restart the application as Administrator.",
                        "Administrator Required");
                }

                _audioService = new AudioService();

                try
                {
                    _deviceChangeNotifier = new DeviceChangeNotifier();
                    _deviceChangeNotifier.DeviceChanged += OnExternalDeviceChanged;
                }
                catch (Exception ex)
                {
                    Log($"Device change notifications unavailable: {ex.Message}");
                }

                RefreshArtTuneKitDetection();
                InitializeTrayIcon();
                UpdateStalePathsOnStartup();
                _isInitialized = true;

                // Bring to front on launch
                this.Activate();
                this.Topmost = _alwaysOnTop;

                // Fetch device data BEFORE clearing the ComboBox to avoid an await gap
                // between clear and repopulate (which causes native WPF render crashes)
                IReadOnlyList<AudioDevice> devices;
                try
                {
                    devices = await _audioService.GetDevicesAsync();
                }
                catch (Exception ex)
                {
                    StyledMessageBox.SafeShowError($"Startup error: {ex.Message}", "Startup Failed");
                    devices = Array.Empty<AudioDevice>();
                }
                _devices = devices.ToList();

                // Now clear + repopulate atomically (no await between these operations)
                DeviceCombo.SelectionChanged -= DeviceCombo_SelectionChanged;
                try
                {
                    DeviceCombo.Items.Clear();

                    if (devices.Count == 0)
                    {
                        DeviceCombo.Items.Add(new AudioDevice { Guid = "no_devices", Name = "No audio devices detected", InterfaceName = "" });
                        DeviceCombo.SelectedIndex = 0;
                        DeviceCombo.IsEnabled = false;
                    }
                    else
                    {
                        DeviceCombo.IsEnabled = true;

                        var firstName = _devices[0].Name ?? string.Empty;
                        if (firstName.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
                            firstName.StartsWith("PS Error", StringComparison.OrdinalIgnoreCase))
                        {
                            StyledMessageBox.ShowWarning(firstName, "Device Load Error");
                        }

                        foreach (var device in devices)
                        {
                            DeviceCombo.Items.Add(device);
                        }
                    }
                }
                finally
                {
                    DeviceCombo.SelectionChanged += DeviceCombo_SelectionChanged;
                    CleanInstallCheck.Checked += CleanInstallCheck_Changed;
                    CleanInstallCheck.Unchecked += CleanInstallCheck_Changed;
                }

                if (_devices.Count > 0)
                {
                    // Try to select last used device
                    if (!string.IsNullOrEmpty(_lastSelectedDeviceGuid))
                    {
                        var lastDevice = _devices.FirstOrDefault(d => d.Guid == _lastSelectedDeviceGuid);
                        if (lastDevice != null)
                        {
                            DeviceCombo.SelectedItem = lastDevice;
                        }
                    }

                    // If no saved device or not found, keep default selection (first device)
                    if (DeviceCombo.SelectedItem == null && _devices.Count > 0)
                    {
                        DeviceCombo.SelectedIndex = 0;
                    }

                    await UpdateIndicatorForSelectedDeviceAsync();
                    await UpdateEapoStatusAsync();
                    UpdateDeviceFormat();
                }

                // Disable controls that require admin if not elevated
                if (!_isAdmin)
                {
                    ReleaseSlider.IsEnabled = false;
                    ReleaseSlider.Opacity = 0.5;
                    ReleaseSlider.ToolTip = "Requires Administrator privileges";

                    LeqPowerButton.IsEnabled = false;
                    LeqPowerButton.Opacity = 0.5;
                    LeqPowerButton.ToolTip = "Requires Administrator privileges";

                    InstallLeqButton.IsEnabled = false;
                    InstallLeqButton.Opacity = 0.5;
                    InstallLeqButton.ToolTip = "Requires Administrator privileges";

                    ResetDeviceButton.IsEnabled = false;
                    ResetDeviceButton.Opacity = 0.5;
                    ResetDeviceButton.ToolTip = "Requires Administrator privileges";

                    Log("Running without admin - LEQ controls disabled.");
                }


                // Finish all UI state updates before revealing the window
                UpdateResetDeviceButtonState();
                UpdateEapoInstalledState();
                UpdateArtTuneGating();

                // Hide loading overlay
                LoadingOverlay.IsHitTestVisible = false;
                LoadingOverlay.Visibility = Visibility.Collapsed;

                // Reveal main window behind the splash, then let WPF
                // finish any final rendering before dismissing the splash
                this.Opacity = 1;
                await Task.Delay(300);

                // Dismiss splash screen
                if (SplashScreen != null)
                {
                    SplashScreen.FadeOutAndClose();
                    SplashScreen = null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[MainWindow] Window_Loaded fatal error: {ex}");

                // Ensure window is visible even on failure — LoadingOverlay will show
                this.Opacity = 1;
                if (SplashScreen != null)
                {
                    SplashScreen.FadeOutAndClose();
                    SplashScreen = null;
                }

                StyledMessageBox.SafeShowError(
                    $"LEQ Control Panel failed to initialize:\n\n{ex.Message}\n\n" +
                    "The application may not function correctly.",
                    "Initialization Error");
            }
        }

        private async Task RefreshDeviceListAsync()
        {
            if (_isRefreshing) return;
            _isRefreshing = true;

            try
            {
                var previousDeviceGuid = GetSelectedDevice()?.Guid;

                // Fetch new device data BEFORE touching the ComboBox.
                // This ensures no await gap between Items.Clear() and repopulation —
                // yielding the UI thread with an empty ComboBox lets WPF render stale
                // visuals, causing native access violations in wpfgfx_cor3.dll.
                var devices = await _audioService.GetDevicesAsync();
                _devices = devices.ToList();

                // Now clear + repopulate atomically (no await between these operations)
                DeviceCombo.SelectionChanged -= DeviceCombo_SelectionChanged;
                try
                {
                    DeviceCombo.Items.Clear();

                    if (devices.Count == 0)
                    {
                        DeviceCombo.Items.Add(new AudioDevice { Guid = "no_devices", Name = "No audio devices detected", InterfaceName = "" });
                        DeviceCombo.SelectedIndex = 0;
                        DeviceCombo.IsEnabled = false;
                    }
                    else
                    {
                        DeviceCombo.IsEnabled = true;

                        foreach (var device in devices)
                        {
                            DeviceCombo.Items.Add(device);
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(previousDeviceGuid))
                    {
                        var matchIndex = _devices.FindIndex(d => d.Guid == previousDeviceGuid);
                        if (matchIndex >= 0 && matchIndex < DeviceCombo.Items.Count)
                        {
                            DeviceCombo.SelectedIndex = matchIndex;
                        }
                        else if (_devices.Count > 0)
                        {
                            DeviceCombo.SelectedIndex = 0;
                        }
                    }
                    else if (_devices.Count > 0)
                    {
                        DeviceCombo.SelectedIndex = 0;
                    }
                }
                finally
                {
                    DeviceCombo.SelectionChanged += DeviceCombo_SelectionChanged;
                }

                if (DeviceCombo.SelectedIndex >= 0)
                {
                    try { await UpdateIndicatorForSelectedDeviceAsync(_devices.AsReadOnly()); }
                    catch (Exception ex) { Debug.WriteLine($"[MainWindow] Refresh: UpdateIndicator failed: {ex.Message}"); }

                    try { await UpdateEapoStatusAsync(); }
                    catch (Exception ex) { Debug.WriteLine($"[MainWindow] Refresh: UpdateEapo failed: {ex.Message}"); }

                    try { UpdateDeviceFormat(); }
                    catch (Exception ex) { Debug.WriteLine($"[MainWindow] Refresh: UpdateDeviceFormat failed: {ex.Message}"); }
                }

                try
                {
                    UpdateEapoInstalledState();
                    RefreshArtTuneKitDetection();
                    UpdateArtTuneGating();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[MainWindow] Refresh: post-refresh UI updates failed: {ex.Message}");
                }
            }
            finally
            {
                _isRefreshing = false;
            }
        }

        private void OnExternalDeviceChanged(object? sender, DeviceChangedEventArgs e)
        {
            if (_isExiting || !_isInitialized || _isManuallyResetting) return;

            Dispatcher.BeginInvoke(() =>
            {
                if (_isExiting || _isManuallyResetting) return;
                _pendingDeviceChanges.Add(e);
                _deviceChangeDebounceTimer.Stop();
                _deviceChangeDebounceTimer.Start();
            });
        }

        private async void DeviceChangeDebounceTimer_Tick(object? sender, EventArgs e)
        {
            _deviceChangeDebounceTimer.Stop();
            if (_isExiting || !_isInitialized || _isManuallyResetting) return;

            // Snapshot and clear pending changes
            var pending = _pendingDeviceChanges.ToList();
            _pendingDeviceChanges.Clear();

            // Snapshot current devices before refresh (for resolving removed device names)
            var previousDevices = _devices?.ToList();

            try
            {
                await RefreshDeviceListAsync();

                // For DefaultDeviceChanged, only keep the last event (final destination,
                // not intermediate bounces as Windows reassigns through devices)
                var lastDefault = pending.LastOrDefault(p => p.ChangeType == DeviceChangeType.DefaultDeviceChanged);
                var filtered = pending
                    .Where(p => p.ChangeType != DeviceChangeType.DefaultDeviceChanged)
                    .ToList();
                if (lastDefault != null) filtered.Add(lastDefault);

                // Deduplicate by (ChangeType, DeviceId) — one line per device per change type
                var unique = filtered
                    .GroupBy(p => (p.ChangeType, DeviceId: p.DeviceId?.ToUpperInvariant() ?? ""))
                    .Select(g => (
                        ChangeType: g.Key.ChangeType,
                        DeviceId: g.First().DeviceId,
                        // For StateChanged, keep the last state value (most recent)
                        LastState: g.LastOrDefault(evt => evt.NewState.HasValue)?.NewState
                    ))
                    .ToList();

                var logged = false;
                foreach (var (changeType, deviceId, lastState) in unique)
                {
                    var device = ResolveDevice(deviceId, previousDevices, _devices);

                    // Skip events for devices we don't track (e.g. capture/microphone endpoints)
                    if (device == null && !string.IsNullOrEmpty(deviceId)) continue;

                    var deviceLabel = FormatDeviceLabel(device);

                    if (changeType == DeviceChangeType.StateChanged)
                    {
                        var stateLabel = lastState switch
                        {
                            1 => "connected",
                            2 => "disabled",
                            4 => "not present",
                            8 => "disconnected",
                            _ => null
                        };
                        var suffix = stateLabel != null ? $" ({stateLabel})" : "";
                        Log(deviceLabel != null ? $"Device{suffix}: {deviceLabel}" : $"Device state changed{suffix}");
                    }
                    else
                    {
                        var label = changeType switch
                        {
                            DeviceChangeType.DeviceAdded => "Device added",
                            DeviceChangeType.DeviceRemoved => "Device removed",
                            DeviceChangeType.DefaultDeviceChanged => "Default device changed",
                            _ => "Device changed"
                        };
                        Log(deviceLabel != null ? $"{label}: {deviceLabel}" : label);
                    }
                    logged = true;
                }

                if (!logged)
                    Log("Audio device change detected");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[MainWindow] DeviceChangeDebounceTimer_Tick error: {ex}");
                Log($"Warning: Failed to refresh device list after device change: {ex.Message}");
            }
        }

        private static AudioDevice? ResolveDevice(string? deviceId, List<AudioDevice>? previousDevices, List<AudioDevice>? currentDevices)
        {
            if (string.IsNullOrEmpty(deviceId)) return null;

            // COM device IDs use format: {0.0.0.00000000}.{guid-here}
            // AudioDevice.Guid is the registry subkey: {guid-here}
            AudioDevice? Find(List<AudioDevice>? devices)
            {
                if (devices == null) return null;
                return devices.FirstOrDefault(d =>
                    !string.IsNullOrEmpty(d.Guid) && deviceId.Contains(d.Guid, StringComparison.OrdinalIgnoreCase));
            }

            return Find(previousDevices) ?? Find(currentDevices);
        }

        private static string? FormatDeviceLabel(AudioDevice? device)
        {
            if (device == null) return null;
            if (!string.IsNullOrEmpty(device.InterfaceName))
                return $"{device.Name} ({device.InterfaceName})";
            return device.Name;
        }

        private async void DeviceCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!_isInitialized || _audioService == null || DeviceCombo.SelectedItem == null)
            {
                return;
            }

            try
            {
                // Hide overlays instantly during device switch
                FadeOverlay(LeqNotDetectedOverlay, false, duration: 0);
                FadeOverlay(LeqLoadingOverlay, false, duration: 0);

                // Now update all UI state with fresh device data
                await UpdateAllUIState();

                // Log device status after data is loaded
                var device = GetSelectedDevice();
                if (device != null)
                {
                    await LogInitializationAsync(device);

                    // Save last selected device
                    try
                    {
                        using var key = Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                        key?.SetValue("LastSelectedDeviceGuid", device.Guid);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[MainWindow] Failed to save device selection to registry: {ex.Message}");
                        // Non-critical - device will work, just won't remember selection on restart
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR: Device selection change failed - {ex.Message}");
                #if DEBUG
                Debug.WriteLine("DeviceCombo_SelectionChanged error: " + ex);
                #endif
            }

            // Update reset device button state based on selected device
            UpdateResetDeviceButtonState();

            // Apply Art Tune device gating if ArtTuneKit is installed
            UpdateArtTuneGating();
        }

        private async void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshButton.IsEnabled = false;
            try
            {
                await RefreshDeviceListAsync();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LEQControlPanel] RefreshButton_Click failed: {ex.Message}");
            }
            finally
            {
                RefreshButton.IsEnabled = true;
            }
        }

        private async void InstallLeqButton_Click(object sender, RoutedEventArgs e)
        {
            var device = GetSelectedDevice();
            if (device is null || string.IsNullOrWhiteSpace(device.Guid))
            {
                return;
            }

            // Handle "Fix Device LEQ" mode — CompositeFX key removal
            if (InstallLeqText.Text == "Fix Device LEQ")
            {
                await FixDeviceCompositeFxAsync(device);
                return;
            }

            // Handle "Fix LEQ Registry" mode — CLSID repair
            if (InstallLeqText.Text == "Fix LEQ Registry")
            {
                await FixClsidFromButtonAsync();
                return;
            }

            try
            {
            // Check if E-APO is configured on this device (used for post-install dialog)
            bool eapoOnDevice = false;
            try
            {
                eapoOnDevice = await _audioService.GetEapoStatusAsync(device.Guid) == true;
            }
            catch (Exception ex)
            {
                Log($"Warning: Could not check E-APO status: {ex.Message}");
            }

            // Show pre-install warning if E-APO is active on this device
            if (eapoOnDevice && !_skipEapoWarning)
            {
                Log("E-APO detected on device - showing conflict warning...");

                var warningDialog = new EapoWarningDialog();
                warningDialog.Owner = this;
                var result = warningDialog.ShowDialog();

                if (result != true || !warningDialog.Proceed)
                {
                    Log("LEQ installation cancelled - user chose not to override E-APO.");
                    return;
                }

                // Save "Don't ask again" preference
                if (warningDialog.DontAskAgain)
                {
                    _skipEapoWarning = true;
                    try
                    {
                        using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                        key?.SetValue("SkipEapoWarning", true);
                        Log("E-APO warning preference saved - won't ask again.");
                    }
                    catch (Exception ex)
                    {
                        Log($"Failed to save E-APO warning preference: {ex.Message}");
                    }
                }

                Log("User chose to proceed with LEQ installation despite E-APO conflict.");
            }

            // If LEQ was previously installed but E-APO displaced it from the chain,
            // offer a clean install to fix it
            if (device.EapoChildBroken)
            {
                Log("E-APO chain broken detected - LEQ was installed but displaced from chain");
                var chainFixDialog = new EapoChainFixDialog(device.InterfaceName ?? device.Name, device.Name);
                chainFixDialog.Owner = this;
                var chainResult = chainFixDialog.ShowDialog();

                if (chainResult != true || !chainFixDialog.Proceed)
                {
                    Log("LEQ installation cancelled - user declined chain fix.");
                    return;
                }

                // Enable clean install for the broken chain fix
                if (CleanInstallCheck != null)
                {
                    CleanInstallCheck.IsChecked = true;
                }
            }

            if (device == null) return; // Redundant guard — satisfies CS8602 after long method body

            Log("Installing LEQ...");
            InstallLeqButton.IsEnabled = false;

            // Show loading overlay with install-specific text
            ShowLeqLoadingOverlay("Installing LEQ...");

            // Suspend COM callbacks before install — the PS script restarts audiosrv internally
            _deviceChangeNotifier?.Suspend();
            try
            {
                bool doClean = CleanInstallCheck?.IsChecked == true;
                var installed = await _audioService.InstallLeqAsync(device.Guid, doClean);
                if (!installed)
                {
                    Log("LEQ installation failed (driver may not support enhancements).");

                    // Check if Clean Install was already tried
                    bool wasCleanInstall = CleanInstallCheck?.IsChecked == true;

                    string message;
                    if (wasCleanInstall)
                    {
                        // Already tried clean install - give full troubleshooting
                        message = "LEQ installation could not complete even with Clean Install.\n\n" +
                            "This usually means:\n" +
                            "\u2022 Your audio driver doesn't support Windows audio enhancements\n" +
                            "\u2022 The device is HDMI/DisplayPort/Optical (these rarely support enhancements)\n" +
                            "\u2022 Windows 11 24H2+ may require manual driver installation\n\n" +
                            "Troubleshooting:\n" +
                            "1. Open Sound Settings and check if your device has an 'Enhancements' tab\n" +
                            "2. If the tab is empty or missing, try installing your motherboard's audio drivers manually\n" +
                            "3. Some USB DACs and gaming headsets don't expose Windows audio enhancements\n\n" +
                            "Would you like to restart audio services?";
                    }
                    else
                    {
                        // Suggest trying Clean Install first
                        message = "LEQ installation could not complete.\n\n" +
                            "\uD83D\uDCA1 TIP: Try checking 'Clean Install' and running again.\n" +
                            "This forces a complete reinstall of the Enhancement tab registry keys.\n\n" +
                            "If Clean Install also fails:\n" +
                            "\u2022 Your audio driver may not support Windows audio enhancements\n" +
                            "\u2022 HDMI/DisplayPort/Optical outputs rarely support enhancements\n" +
                            "\u2022 On Windows 11 24H2+, try installing motherboard audio drivers manually\n\n" +
                            "Would you like to restart audio services?";
                    }

                    var failResult = StyledMessageBox.ShowYesNo(message, "LEQ Installation Failed");

                    if (failResult == MessageBoxResult.Yes)
                    {
                        var (ok, reason) = await RestartAudioServiceWithSuspendAsync();
                        Log(ok ? "Audio services restarted after failed installation."
                               : $"Audio service restart failed after failed installation: {reason}");
                    }
                    UpdateLeqIndicator(false);
                    return;
                }

                // Refresh device data to pick up new LEQ state
                await RefreshDeviceListAsync();

                try
                {
                    SoundPanelHelper.OpenSoundPanel();
                    Log("Opened Windows Sound Panel for LEQ verification.");
                }
                catch (Exception ex)
                {
                    Log($"Unable to open Sound Panel for verification: {ex.Message}");
                }

                var verifyDialog = new LeqVerifyDialog { Owner = this };
                verifyDialog.ShowDialog();
                var verifyResult = verifyDialog.Result;
                if (verifyResult == MessageBoxResult.Yes)
                {
                    Log("LEQ installation verified.");
                    UpdateLeqIndicator(true, "Verified");
                    UpdateInstallButtonState(true); // Update button to "LEQ Installed" state

                    // If E-APO was on this device, prompt user to reconfigure it
                    if (eapoOnDevice)
                    {
                        Log("E-APO was active - prompting user to run Device Selector...");
                        var eapoResult = StyledMessageBox.ShowYesNo(
                            "E-APO needs to be reconfigured on this device.\n\n" +
                            "Open Device Selector, make sure Art Tune is checked, and close it.\n\n" +
                            "Open Device Selector now?",
                            "Restore E-APO");
                        if (eapoResult == MessageBoxResult.Yes)
                        {
                            LaunchEapoDeviceSelector();
                        }
                    }
                }
                else
                {
                    Log("LEQ verification failed.");

                    bool wasCleanInstall = CleanInstallCheck?.IsChecked == true;

                    string message;
                    if (wasCleanInstall)
                    {
                        message = "LEQ verification was unsuccessful even with Clean Install.\n\n" +
                            "The Enhancement tab should be visible but LEQ isn't working.\n\n" +
                            "Please check:\n" +
                            "\u2022 Is 'Loudness Equalization' visible and checked in the Enhancements tab?\n" +
                            "\u2022 Did the test bars go full green repeatedly when clicking 'Test'?\n\n" +
                            "If the checkbox exists but has no effect:\n" +
                            "\u2022 This is a known issue on Windows 11 24H2\n" +
                            "\u2022 Try installing older/different audio drivers\n" +
                            "\u2022 Some driver versions remove LEQ support entirely\n\n" +
                            "Would you like to restart audio services?";
                    }
                    else
                    {
                        message = "LEQ verification was unsuccessful.\n\n" +
                            "\uD83D\uDCA1 TIP: Try checking 'Clean Install' and running again.\n\n" +
                            "Please check:\n" +
                            "\u2022 Is 'Loudness Equalization' visible and checked in the Enhancements tab?\n" +
                            "\u2022 Did the test bars go full green repeatedly when clicking 'Test'?\n\n" +
                            "If Enhancements tab is empty:\n" +
                            "Your audio driver doesn't support Windows enhancements.\n\n" +
                            "Would you like to restart audio services?";
                    }

                    var revertResult = StyledMessageBox.ShowYesNo(message, "LEQ Verification Failed");

                    if (revertResult == MessageBoxResult.Yes)
                    {
                        var (ok, reason) = await RestartAudioServiceWithSuspendAsync();
                        Log(ok ? "Audio services restarted after failed verification."
                               : $"Audio service restart failed after failed verification: {reason}");
                    }
                    UpdateLeqIndicator(false);
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR: LEQ installation failed - {ex.Message}");
                StyledMessageBox.SafeShowError(
                    $"LEQ installation encountered an error:\n\n{ex.Message}\n\n" +
                    "Try restarting the application or check the console for details.",
                    "Installation Error");
                UpdateLeqIndicator(false);
            }
            finally
            {
                // Resume COM callbacks after install (suspend was before InstallLeqAsync)
                ResumeDeviceNotifierAndRefresh();

                // Hide loading overlay
                HideLeqLoadingOverlay();

                InstallLeqButton.IsEnabled = true;
            }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[MainWindow] InstallLeqButton_Click error: {ex}");
                Log($"ERROR: LEQ install action failed - {ex.Message}");
            }
            finally
            {
                // Ensure overlay is always hidden, even if an exception occurred before the inner try
                if (LeqLoadingOverlay != null && LeqLoadingOverlay.Opacity > 0)
                {
                    LeqLoadingOverlay.Opacity = 0;
                    LeqLoadingOverlay.IsHitTestVisible = false;
                }
                InstallLeqButton.IsEnabled = true;
            }
        }

        private void SoundSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            Log("Opening Windows Sound settings...");
            try
            {
                SoundPanelHelper.OpenSoundPanel();
                Log("Windows Sound settings opened.");
            }
            catch (Exception ex)
            {
                Log($"Failed to open Windows Sound settings: {ex.Message}");
                StyledMessageBox.SafeShowError($"Unable to open Windows Sound settings:\n\n{ex.Message}", "Sound Settings");
            }
        }

        /// <summary>
        /// Suspend COM callbacks, restart audio service, wait for stabilization, resume.
        /// Centralizes the suspend/restart/resume pattern so every restart path is safe.
        /// </summary>
        private async Task<(bool success, string? reason)> RestartAudioServiceWithSuspendAsync()
        {
            _deviceChangeNotifier?.Suspend();
            try
            {
                var result = await _audioService.RestartAudioServiceAsync();
                await Task.Delay(500); // let the new audio service stabilize
                return result;
            }
            finally
            {
                ResumeDeviceNotifierAndRefresh();
            }
        }

        private async void RestartAudioButton_Click(object sender, RoutedEventArgs e)
        {
            RestartAudioButton.IsEnabled = false;
            Log("Restarting audio service...");
            try
            {
                var (restarted, reason) = await RestartAudioServiceWithSuspendAsync();
                if (restarted)
                    StyledMessageBox.ShowInfo("Audio service restarted successfully.", "Audio Service");
                else
                    StyledMessageBox.ShowWarning($"Unable to restart the Windows audio service.\n\n{reason}", "Audio Service");
            }
            catch (Exception ex)
            {
                Log($"ERROR: Audio service restart failed - {ex.Message}");
                StyledMessageBox.SafeShowError($"Audio service restart failed:\n\n{ex.Message}", "Audio Service");
            }
            finally
            {
                RestartAudioButton.IsEnabled = true;
                UpdateEapoInstalledState();
                UpdateArtTuneGating();
            }
        }

        private static bool CheckEapoInstalled()
        {
            return File.Exists(@"C:\Program Files\EqualizerAPO\Editor.exe");
        }

        private void UpdateEapoInstalledState()
        {
            _isEapoInstalled = CheckEapoInstalled();

            if (_isEapoInstalled)
            {
                EapoDeviceSelectorButton.Style = (Style)FindResource("InstallLeqButtonStyle");
                EapoButtonIcon.Text = "\xE9E9";
                EapoButtonIcon.FontFamily = new System.Windows.Media.FontFamily("Segoe MDL2 Assets");
                EapoButtonText.Text = "E-APO Device Selector";
                EapoDeviceSelectorButton.ToolTip = "Open E-APO Device Selector";
            }
            else
            {
                EapoDeviceSelectorButton.Style = (Style)FindResource("NeutralButtonStyle");
                EapoButtonIcon.Text = "\u2B07";
                EapoButtonIcon.FontFamily = new System.Windows.Media.FontFamily("Segoe UI");
                EapoButtonIcon.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#22AA22"));
                EapoButtonText.Text = "Get E-APO";
                EapoDeviceSelectorButton.ToolTip = "Download Equalizer APO";
            }
        }

        private void LaunchDeviceSelector_Click(object sender, RoutedEventArgs e)
        {
            if (_isEapoInstalled)
            {
                LaunchEapoDeviceSelector();
            }
            else
            {
                Process.Start(new ProcessStartInfo("https://sourceforge.net/projects/equalizerapo/") { UseShellExecute = true })?.Dispose();
            }
        }

        /// <summary>
        /// Launches the E-APO Device Selector executable directly.
        /// Returns true if launched successfully, false otherwise.
        /// </summary>
        internal static bool LaunchEapoDeviceSelector()
        {
            var exePath = @"C:\Program Files\EqualizerAPO\DeviceSelector.exe";
            if (File.Exists(exePath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo(exePath)
                    {
                        UseShellExecute = true,
                        WorkingDirectory = Path.GetDirectoryName(exePath)!
                    })?.Dispose();
                    return true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[MainWindow] Failed to launch DeviceSelector: {ex}");
                    StyledMessageBox.SafeShowError(
                        $"Could not launch Device Selector:\n\n{ex.Message}",
                        "Device Selector");
                    return false;
                }
            }

            StyledMessageBox.ShowInfo(
                "Couldn't find DeviceSelector.exe.\n\nExpected location:\nC:\\Program Files\\EqualizerAPO\\DeviceSelector.exe",
                "Device Selector");
            return false;
        }

        private void UpdateResetDeviceButtonState()
        {
            if (!_isAdmin)
            {
                ResetDeviceButton.IsEnabled = false;
                ResetDeviceButton.Opacity = 0.5;
                ResetDeviceButton.ToolTip = "Requires Administrator privileges";
                return;
            }

            var device = GetSelectedDevice();
            if (device?.InterfaceName != null &&
                device.InterfaceName.StartsWith("Voicemeeter", StringComparison.OrdinalIgnoreCase))
            {
                ResetDeviceButton.IsEnabled = false;
                ResetDeviceButton.Opacity = 0.5;
                ResetDeviceButton.ToolTip = "Cannot reset Voicemeeter devices";
                return;
            }

            ResetDeviceButton.IsEnabled = true;
            ResetDeviceButton.Opacity = 1.0;
            ResetDeviceButton.ToolTip = null;
        }

        private async void ResetDeviceButton_Click(object sender, RoutedEventArgs e)
        {
            var device = GetSelectedDevice();
            if (device is null || string.IsNullOrWhiteSpace(device.Guid))
            {
                StyledMessageBox.ShowWarning("No device selected.", "Reset Device");
                return;
            }

            // Guard: Voicemeeter devices should already be disabled
            if (device.InterfaceName != null &&
                device.InterfaceName.StartsWith("Voicemeeter", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            // Determine confirmation style based on device type
            MessageBoxResult confirmResult;
            var isVirtualCable = device.InterfaceName != null &&
                (device.InterfaceName.StartsWith("VB-Audio", StringComparison.OrdinalIgnoreCase) ||
                 device.InterfaceName.StartsWith("Hi-Fi", StringComparison.OrdinalIgnoreCase));

            if (isVirtualCable)
            {
                confirmResult = StyledMessageBox.ShowDanger(
                    "This will completely remove the virtual audio cable driver from your system. " +
                    "You will need to reinstall it afterward.\n\n" +
                    "Are you sure?",
                    "Remove Virtual Audio Cable?",
                    "Remove",
                    "Cancel");
            }
            else
            {
                if (!_skipResetDeviceWarning)
                {
                    var warningDialog = new Dialogs.ResetDeviceWarningDialog();
                    warningDialog.Owner = this;
                    var result = warningDialog.ShowDialog();

                    if (result != true || !warningDialog.Proceed)
                        return;

                    if (warningDialog.DontShowAgain)
                    {
                        _skipResetDeviceWarning = true;
                        try
                        {
                            using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                            key?.SetValue("SkipResetDeviceWarning", true);
                            Log("Reset device warning preference saved - won't show again.");
                        }
                        catch (Exception ex) { Log($"Failed to save reset warning preference: {ex.Message}"); }
                    }
                }

                confirmResult = MessageBoxResult.Yes;
            }

            if (confirmResult != MessageBoxResult.Yes)
            {
                return;
            }

            ResetDeviceButton.IsEnabled = false;
            _isManuallyResetting = true;
            _deviceChangeDebounceTimer.Stop();
            Log($"Resetting device: {device.Name}...");

            try
            {
            // Show loading overlay on LEQ controls
            ShowLeqLoadingOverlay("Resetting device...");

            // Show pulsing loading bar under device dropdown
            StartDeviceLoadingBar();

            // Suspend COM callbacks — pnputil device operations trigger device state changes
            _deviceChangeNotifier?.Suspend();
            try
            {
                var (success, message) = await _audioService.ResetDeviceAsync(device.Guid, device.Name);

                if (success)
                {
                    if (isVirtualCable)
                    {
                        // Virtual cables are removed entirely — wait for disappearance, not reappearance
                        Log("Virtual cable removal initiated. Waiting for device to disappear...");
                        if (LoadingOverlayText != null)
                            LoadingOverlayText.Text = "Removing virtual cable...";

                        bool deviceGone = false;

                        // Fast phase: registry polling every 500ms for 5s
                        var removeSw = Stopwatch.StartNew();
                        while (removeSw.ElapsedMilliseconds < 5000)
                        {
                            await Task.Delay(500);
                            int state = AudioService.ReadDeviceStateFromRegistry(device.Guid);
                            if (state == 0) // Registry key gone
                            {
                                deviceGone = true;
                                Log($"Virtual cable disappeared after {removeSw.ElapsedMilliseconds}ms (fast path).");
                                break;
                            }
                            Log($"Waiting for virtual cable to disappear... ({removeSw.ElapsedMilliseconds}ms)");
                        }

                        // Slow phase: full enumeration check
                        if (!deviceGone)
                        {
                            for (int i = 0; i < 5; i++)
                            {
                                await Task.Delay(2000);
                                var allDevices = await _audioService.GetDevicesAsync(includeInactive: true);
                                bool stillPresent = allDevices.Any(d =>
                                    string.Equals(d.Guid, device.Guid, StringComparison.OrdinalIgnoreCase));
                                if (!stillPresent)
                                {
                                    deviceGone = true;
                                    Log("Virtual cable confirmed removed.");
                                    break;
                                }
                                Log($"Waiting for virtual cable to disappear... ({i + 1}/5)");
                            }
                        }

                        // Refresh the UI device list
                        try
                        {
                            await RefreshDeviceListAsync();
                        }
                        catch (Exception refreshEx)
                        {
                            Log($"Warning: Failed to refresh device list after removal: {refreshEx.Message}");
                        }

                        await Task.Delay(1500);

                        if (deviceGone)
                        {
                            StyledMessageBox.ShowInfo(
                                "Virtual audio cable removed successfully.",
                                "Reset Device");
                        }
                        else
                        {
                            Log("Virtual cable may still be present after polling.");
                            StyledMessageBox.ShowWarning(
                                "The virtual audio cable may not have been fully removed.\n\n" +
                                "A reboot may be required to complete the removal.",
                                "Reset Device");
                        }
                    }
                    else
                    {
                    // Physical device — wait for reappearance
                    Log("Device reset completed. Waiting for device to re-enumerate...");
                    if (LoadingOverlayText != null)
                    {
                        LoadingOverlayText.Text = "Waiting for device...";
                    }

                    // Phase 1: Fast registry-based polling (first 5s, every 500ms)
                    bool deviceFound = false;
                    var reEnumSw = Stopwatch.StartNew();
                    while (reEnumSw.ElapsedMilliseconds < 5000)
                    {
                        await Task.Delay(500);
                        int state = AudioService.ReadDeviceStateFromRegistry(device.Guid);
                        if (state > 0) // Key exists — device has re-enumerated
                        {
                            // Confirm with full PowerShell enumeration
                            var allDevices = await _audioService.GetDevicesAsync(includeInactive: true);
                            var reappeared = allDevices.Any(d =>
                                string.Equals(d.Guid, device.Guid, StringComparison.OrdinalIgnoreCase));
                            if (reappeared)
                            {
                                deviceFound = true;
                                Log($"Device re-detected after {reEnumSw.ElapsedMilliseconds}ms (fast path).");
                                break;
                            }
                        }
                        Log($"Fast polling for device... ({reEnumSw.ElapsedMilliseconds}ms)");
                    }

                    // Phase 2: Standard progressive backoff if fast path didn't find it
                    if (!deviceFound)
                    {
                        int remainingAttempts = 10; // ~35s: 5×3s + 5×5s
                        for (int i = 0; i < remainingAttempts; i++)
                        {
                            int delayMs = i < 5 ? 3000 : 5000;
                            await Task.Delay(delayMs);

                            var allDevices = await _audioService.GetDevicesAsync(includeInactive: true);
                            var reappeared = allDevices.Any(d =>
                                string.Equals(d.Guid, device.Guid, StringComparison.OrdinalIgnoreCase));
                            if (reappeared)
                            {
                                deviceFound = true;
                                Log("Device re-detected successfully.");
                                break;
                            }
                            Log($"Waiting for device to reappear... ({i + 1}/{remainingAttempts})");
                        }
                    }

                    // Allow Windows PnP to finish setting up audio endpoints before refreshing
                    if (deviceFound)
                    {
                        Log("Device found - waiting for audio stack to settle...");
                        if (LoadingOverlayText != null)
                            LoadingOverlayText.Text = "Device settling...";

                        // Poll for DeviceState == 1 (Ready) instead of fixed 2s wait
                        var settleSw = Stopwatch.StartNew();
                        while (settleSw.ElapsedMilliseconds < 3000)
                        {
                            int state = AudioService.ReadDeviceStateFromRegistry(device.Guid);
                            if (state == 1)
                            {
                                Log($"Device audio stack ready after {settleSw.ElapsedMilliseconds}ms");
                                break;
                            }
                            await Task.Delay(200);
                        }
                    }

                    // Refresh the UI device list now that re-detection is complete
                    try
                    {
                        await RefreshDeviceListAsync();
                    }
                    catch (Exception refreshEx)
                    {
                        Log($"Warning: Failed to refresh device list after reset: {refreshEx.Message}");
                    }

                    // Suppress COM notifications a bit longer so the debounce timer
                    // doesn't race with the just-completed refresh
                    await Task.Delay(1500);

                    if (deviceFound)
                    {
                        StyledMessageBox.ShowInfo(
                            "Device reset completed successfully.",
                            "Reset Device");
                    }
                    else
                    {
                        Log("Device did not reappear after polling.");

                        // Show recovery dialog loop — let user manually retry
                        while (true)
                        {
                            var retryResult = StyledMessageBox.ShowYesNo(
                                "Your device was reset but hasn't reappeared yet.\n\n" +
                                "Please unplug your device, wait 5 seconds, then plug it back in.",
                                "Device Didn't Reappear",
                                "Refresh Device List",
                                "Close");

                            if (retryResult != MessageBoxResult.Yes)
                                break;

                            try
                            {
                                await RefreshDeviceListAsync();
                            }
                            catch (Exception refreshEx)
                            {
                                Log($"Warning: Failed to refresh device list: {refreshEx.Message}");
                            }

                            var refreshedDevices = await _audioService.GetDevicesAsync(includeInactive: true);
                            if (refreshedDevices.Any(d =>
                                string.Equals(d.Guid, device.Guid, StringComparison.OrdinalIgnoreCase)))
                            {
                                StyledMessageBox.ShowInfo(
                                    "Device reset completed successfully.",
                                    "Reset Device");
                                break;
                            }
                        }
                    }
                    }
                }
                else
                {
                    Log($"Device reset failed: {message}");
                    StyledMessageBox.SafeShowError(
                        $"Device reset failed:\n\n{message}",
                        "Reset Device");
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR: Device reset failed - {ex.Message}");
                StyledMessageBox.SafeShowError(
                    $"Device reset encountered an error:\n\n{ex.Message}",
                    "Reset Device");
            }
            finally
            {
                // Resume COM callbacks after device reset
                ResumeDeviceNotifierAndRefresh();

                // Hide loading overlay
                HideLeqLoadingOverlay();

                // Hide pulsing loading bar
                StopDeviceLoadingBar();
            }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[MainWindow] ResetDeviceButton_Click error: {ex}");
                Log($"ERROR: Device reset action failed - {ex.Message}");
            }
            finally
            {
                // Ensure overlay is always hidden, even if an exception occurred before the inner try
                if (LeqLoadingOverlay != null && LeqLoadingOverlay.Opacity > 0)
                {
                    LeqLoadingOverlay.Opacity = 0;
                    LeqLoadingOverlay.IsHitTestVisible = false;
                }

                _isManuallyResetting = false;
                ResetDeviceButton.IsEnabled = true;
                UpdateResetDeviceButtonState();
            }
        }

        private void InfoButton_Click(object sender, RoutedEventArgs e)
        {
            Log("Opened About panel.");
            var aboutWindow = new AboutWindow();
            aboutWindow.Owner = this;
            aboutWindow.ShowDialog();
        }

        private async void LeqPowerButton_Click(object sender, RoutedEventArgs e)
        {
            var device = GetSelectedDevice();
            if (device == null)
            {
                // Revert the automatic toggle if no device selected
                if (sender is ToggleButton btn)
                {
                    btn.IsChecked = !btn.IsChecked;
                }
                return;
            }

            // Store the current UI state (which was already toggled by the button)
            var targetState = LeqPowerButton.IsChecked;

            // Disable the button during the operation
            LeqPowerButton.IsEnabled = false;

            // Show loading overlay with toggle-specific text
            ShowLeqLoadingOverlay("Toggling LEQ...");

            // Suspend COM callbacks — registry writes to FxProperties can trigger device change events
            _deviceChangeNotifier?.Suspend();
            try
            {
                Log("Toggling LEQ...");

                // Call the toggle service
                var result = await _audioService.ToggleLeqAsync(device.Guid);

                if (result == null)
                {
                    // Toggle failed - revert the UI
                    Log("LEQ toggle failed - reverting UI state");
                    LeqPowerButton.IsChecked = !targetState;
                }
                else
                {
                    // Toggle succeeded - update UI to match actual state
                    LeqPowerButton.IsChecked = result.Value;

                    // Poll registry until the written state is readable (typically instant)
                    var toggleSw = Stopwatch.StartNew();
                    while (toggleSw.ElapsedMilliseconds < 300)
                    {
                        var regState = AudioService.ReadLeqStateFromRegistry(device.Guid);
                        if (regState == result.Value)
                            break;
                        await Task.Delay(50);
                    }

                    await UpdateIndicatorForSelectedDeviceAsync();
                    await UpdateEapoStatusAsync();

                    Log(result.Value ? "LEQ enabled." : "LEQ disabled.");
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR: LEQ toggle exception - {ex.Message}");
                // Revert UI on error
                LeqPowerButton.IsChecked = !targetState;
                StyledMessageBox.SafeShowError($"Failed to toggle LEQ:\n\n{ex.Message}", "LEQ Toggle Error");
            }
            finally
            {
                // Resume COM callbacks after toggle
                ResumeDeviceNotifierAndRefresh();

                // Hide loading overlay
                HideLeqLoadingOverlay();

                LeqPowerButton.IsEnabled = true;
            }
        }

        private void LeqPowerButton_Checked(object sender, RoutedEventArgs e)
        {
            // Handle the checked state change - update UI elements
            UpdateLeqIndicator(true);

            // Update readout to show actual values when LEQ is turned ON
            if (ReleaseSlider != null)
            {
                UpdateReleaseVisuals((int)ReleaseSlider.Value);
            }
        }

        private void LeqPowerButton_Unchecked(object sender, RoutedEventArgs e)
        {
            // Handle the unchecked state change - update UI elements
            UpdateLeqIndicator(false);

            // Update readout to show "-" when LEQ is turned OFF
            if (ReleaseSlider != null)
            {
                UpdateReleaseVisuals((int)ReleaseSlider.Value);
            }
        }

        private void UpdateLeqIndicator(bool? state, string? overrideText = null)
        {
            // LEQ status is now indicated by the power button itself
            // LeqStatusDot and LeqStatusState elements were removed from XAML
            SetLeqPowerButtonState(state);
            UpdateToggleLeqMenuState(state);
            UpdateTrayTooltip(state);

            // NEW: Lock the Slider and Readout if LEQ is OFF
            bool isLeqOn = state == true;

            if (ReleaseSlider != null)
            {
                ReleaseSlider.IsEnabled = isLeqOn;
                ReleaseSlider.Opacity = isLeqOn ? 1.0 : 0.3; // Dim it visually
            }

            // Dim the Digital Readout too
            // Readout opacity now handled by slider selection range
            // No need to manually control opacity
        }

        private void UpdateEapoBadge(bool? isActive)
        {
            var statusText = isActive switch
            {
                true => "Active",
                false => "Inactive",
                null => "Not Installed"
            };

            // Removed EapoStatusDot ellipse - coloring now handled by XAML DataTrigger
            if (EapoStatusState != null)
            {
                EapoStatusState.Text = statusText;
                // EapoStatusState shows status text alongside the colored dot
            }
        }

        private void SetLeqPowerButtonState(bool? state)
        {
            if (LeqPowerButton == null)
            {
                return;
            }

            LeqPowerButton.IsChecked = state == true;
            // Note: IsEnabled is now controlled by UpdateLeqToggleAvailability based on HasReleaseTimeKey
        }

        private void UpdateLeqToggleAvailability(AudioDevice? device)
        {
            if (LeqPowerButton == null) return;

            if (device == null)
            {
                LeqPowerButton.IsEnabled = false;
                LeqPowerButton.Opacity = 0.5;
                LeqPowerButton.ToolTip = "No device selected";
            }
            else if (_clsidsBroken)
            {
                // CLSID registrations broken (global — affects all devices)
                LeqPowerButton.IsEnabled = false;
                LeqPowerButton.Opacity = 0.5;
                LeqPowerButton.ToolTip = "CLSID registrations broken - click 'Fix LEQ Registry' to repair";
            }
            else if (!device.LeqConfigured)
            {
                // Device doesn't have LEQ - disable toggle, show install
                LeqPowerButton.IsEnabled = false;
                LeqPowerButton.Opacity = 0.5;
                LeqPowerButton.ToolTip = "LEQ not installed - use Install LEQ button";
            }
            else if (device.HasCompositeFx)
            {
                // LEQ installed but CompositeFX keys blocking (per-device)
                LeqPowerButton.IsEnabled = false;
                LeqPowerButton.Opacity = 0.5;
                LeqPowerButton.ToolTip = "CompositeFX keys blocking LEQ - click 'Fix Device LEQ' to repair";
            }
            else
            {
                // Device has LEQ - enable toggle
                LeqPowerButton.IsEnabled = true;
                LeqPowerButton.Opacity = 1.0;

                // Set tooltip based on E-APO status
                if (device.EapoChildBroken)
                {
                    LeqPowerButton.ToolTip = "⚠️ E-APO chain broken - LEQ toggle won't affect audio.\n\n" +
                        "Fix:\n1. Click 'Clean Install' checkbox\n2. Click 'Install LEQ'\n3. Re-run E-APO Configurator on this device";
                }
                else if (device.EapoStatus == "Active")
                {
                    LeqPowerButton.ToolTip = "E-APO detected. If LEQ toggle doesn't affect audio:\n" +
                        "1. Use 'Clean Install' to reinstall LEQ\n" +
                        "2. Re-run E-APO Configurator on this device\n" +
                        "This re-registers LEQ in E-APO's audio chain.";
                }
                else
                {
                    LeqPowerButton.ToolTip = "Toggle Loudness Equalization";
                }
            }

            // Update Install button state (inverse of toggle availability)
            UpdateInstallButtonState(device?.LeqConfigured ?? false, device?.HasLfxGfx ?? false, device?.HasCompositeFx ?? false, _clsidsBroken);
        }

        private void CleanInstallCheck_Changed(object sender, RoutedEventArgs e)
        {
            // Re-evaluate install button state with current device
            if (DeviceCombo.SelectedItem is AudioDevice device)
            {
                UpdateInstallButtonState(device.LeqConfigured, device.HasLfxGfx, device.HasCompositeFx, _clsidsBroken);
            }
        }

        private void UpdateInstallButtonState(bool leqConfigured, bool hasLfxGfx = true, bool hasCompositeFx = false, bool clsidsBroken = false)
        {
            if (InstallLeqButton == null || InstallLeqIcon == null || InstallLeqText == null || LeqNotDetectedOverlay == null) return;

            bool cleanInstallChecked = CleanInstallCheck.IsChecked == true;

            // 1. Set Clean Install checkbox state
            CleanInstallCheck.IsEnabled = leqConfigured;
            CleanInstallCheck.Opacity = leqConfigured ? 1.0 : 0.5;

            // 2. Reset text foreground (may have been overridden by red state)
            var defaultTextBrush = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1A1A1A"));
            InstallLeqIcon.Foreground = defaultTextBrush;
            InstallLeqText.Foreground = defaultTextBrush;

            // 3. Set button text/icon and style based on state
            if (clsidsBroken && !cleanInstallChecked)
            {
                // CLSIDs broken — offer registry fix (red button, global — regardless of device state)
                InstallLeqButton.Style = (Style)FindResource("GlowingRedInstallButtonStyle");
                InstallLeqButton.IsEnabled = true;
                InstallLeqIcon.Text = "\uE90F"; // Wrench icon
                InstallLeqText.Text = "Fix LEQ Registry";
                var whiteBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.White);
                InstallLeqIcon.Foreground = whiteBrush;
                InstallLeqText.Foreground = whiteBrush;
                InstallLeqButton.ToolTip = "Fix broken CLSID registrations for LEQ";
            }
            else if (leqConfigured && hasCompositeFx && !cleanInstallChecked)
            {
                // LEQ installed but CompositeFX keys blocking — offer fix (red button, per-device)
                InstallLeqButton.Style = (Style)FindResource("GlowingRedInstallButtonStyle");
                InstallLeqButton.IsEnabled = true;
                InstallLeqIcon.Text = "\uE90F"; // Wrench icon (Segoe MDL2 Assets)
                InstallLeqText.Text = "Fix Device LEQ";
                var whiteBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.White);
                InstallLeqIcon.Foreground = whiteBrush;
                InstallLeqText.Foreground = whiteBrush;
                InstallLeqButton.ToolTip = "Remove CompositeFX keys blocking LEQ on this device";
            }
            else if (leqConfigured && !cleanInstallChecked)
            {
                InstallLeqButton.Style = (Style)FindResource("InstallLeqButtonStyle");
                InstallLeqButton.IsEnabled = false;
                InstallLeqIcon.Text = "\u2713";
                InstallLeqText.Text = "LEQ Installed";
                InstallLeqButton.ToolTip = "Loudness Equalization is installed";
                InstallLeqIcon.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#22AA22"));
                var greyBrush = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#999999"));
                InstallLeqText.Foreground = greyBrush;
            }
            else if (leqConfigured && cleanInstallChecked)
            {
                InstallLeqButton.Style = (Style)FindResource("GlowingInstallButtonStyle");
                InstallLeqButton.IsEnabled = true;
                InstallLeqIcon.Text = "\u21BB";
                InstallLeqText.Text = "Reinstall LEQ";
                InstallLeqButton.ToolTip = "Reinstall Loudness Equalization";
            }
            else if (!hasLfxGfx)
            {
                InstallLeqButton.Style = (Style)FindResource("InstallLeqButtonStyle");
                InstallLeqButton.IsEnabled = false;
                InstallLeqIcon.Text = "\u2B07";
                InstallLeqText.Text = "Install LEQ";
                InstallLeqButton.ToolTip = "Run E-APO Device Selector first";
            }
            else
            {
                InstallLeqButton.Style = (Style)FindResource("GlowingInstallButtonStyle");
                InstallLeqButton.IsEnabled = true;
                InstallLeqIcon.Text = "\u2B07";
                InstallLeqText.Text = "Install LEQ";
                InstallLeqButton.ToolTip = "Install Loudness Equalization";
                InstallLeqIcon.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#22AA22"));
            }

            // 4. Set overlay state and message
            bool showOverlay;
            if (clsidsBroken)
            {
                LeqOverlayTitle.Text = "CLSID Registrations Broken";
                LeqOverlaySubtitle.Text = "Click 'Fix LEQ Registry' button below to repair";
                showOverlay = true;
            }
            else if (leqConfigured && hasCompositeFx)
            {
                LeqOverlayTitle.Text = "CompositeFX Keys Blocking LEQ";
                LeqOverlaySubtitle.Text = "Click 'Fix Device LEQ' button below to repair";
                showOverlay = true;
            }
            else if (!leqConfigured && !hasLfxGfx)
            {
                LeqOverlayTitle.Text = "Device Not Ready for LEQ";
                LeqOverlaySubtitle.Text = "Configure LFX/GFX in E-APO Device Selector first";
                showOverlay = true;
            }
            else if (!leqConfigured)
            {
                LeqOverlayTitle.Text = "LEQ Enhancement Not Detected";
                LeqOverlaySubtitle.Text = "Click 'Install LEQ' button below";
                showOverlay = true;
            }
            else
            {
                showOverlay = false;
            }
            FadeOverlay(LeqNotDetectedOverlay, showOverlay);
        }

        private void UpdateToggleLeqMenuState(bool? state)
        {
            if (_toggleLeqMenuItem == null)
            {
                return;
            }

            _toggleLeqMenuItem.Checked = state == true;
            _toggleLeqMenuItem.Text = state switch
            {
                true => "Toggle LEQ \u2713",
                false => "Toggle LEQ \u2715",
                null => "Toggle LEQ"
            };
        }

        private void UpdateTrayTooltip(bool? leqState = null)
        {
            if (_trayIcon == null)
            {
                return;
            }

            if (leqState.HasValue)
            {
                _lastLeqState = leqState;
            }

            var device = GetSelectedDevice();
            var deviceName = string.IsNullOrWhiteSpace(device?.Name) ? "No Device" : device.Name;
            var stateText = _lastLeqState switch
            {
                true => "On",
                false => "Off",
                null => "Unknown"
            };

            var text = $"LEQ Control Panel\n{deviceName}: LEQ {stateText}";
            if (text.Length > 63)
            {
                text = text.Substring(0, 63);
            }

            _trayIcon.Text = text;
        }

        private static ScrollViewer? GetScrollViewer(DependencyObject element)
        {
            if (element is ScrollViewer scrollViewer)
                return scrollViewer;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(element); i++)
            {
                var child = VisualTreeHelper.GetChild(element, i);
                var result = GetScrollViewer(child);
                if (result != null)
                    return result;
            }
            return null;
        }

        private static readonly Brush WarnBrush = new SolidColorBrush(Color.FromRgb(0xA0, 0x88, 0x30)); // dim amber
        private static readonly Brush ErrorBrush = new SolidColorBrush(Color.FromRgb(0xA0, 0x50, 0x50)); // dim red

        static MainWindow()
        {
            WarnBrush.Freeze();
            ErrorBrush.Freeze();
        }

        private void Log(string message)
        {
            if (ConsoleList == null) return;

            // Ensure we're on UI thread
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => Log(message));
                return;
            }

            var entry = $"[{DateTime.Now:HH:mm}] {message}";
            var textBlock = new System.Windows.Controls.TextBlock
            {
                Text = entry,
                TextWrapping = TextWrapping.Wrap
            };

            // Color warnings and errors
            if (message.Contains("[WARN]", StringComparison.OrdinalIgnoreCase) ||
                message.StartsWith("Warning:", StringComparison.OrdinalIgnoreCase))
            {
                textBlock.Foreground = WarnBrush;
            }
            else if (message.Contains("ERROR", StringComparison.OrdinalIgnoreCase))
            {
                textBlock.Foreground = ErrorBrush;
            }

            ConsoleList.Items.Add(textBlock);

            // Get ScrollViewer and scroll to end - this is the reliable way
            Dispatcher.BeginInvoke(DispatcherPriority.Loaded, new Action(() =>
            {
                var scrollViewer = GetScrollViewer(ConsoleList);
                scrollViewer?.ScrollToEnd();
            }));
        }

        private async Task LogInitializationAsync(AudioDevice device)
        {
            try
            {
                var leqState = await _audioService.GetLeqStateAsync(device.Guid);
                var eapoStatus = await _audioService.GetEapoStatusAsync(device.Guid);

                string leqStatus = leqState == true ? "ACTIVE" : "INACTIVE";
                string eapoText = eapoStatus == true ? "ACTIVE" : "INACTIVE";
                string speakerConfig = GetSpeakerConfigName(device.Channels);

                Log($"[INIT] DEVICE: {device.Name}");
                Log($"[STAT] LEQ: {leqStatus} | RT: {device.ReleaseTime} | EAPO: {eapoText} | LFX/GFX: {(device.HasLfxGfx ? "YES" : "NO")} | FORMAT: {speakerConfig} / {device.BitDepth}-bit {device.SampleRate}Hz");
                Log("[SYS]  SYSTEM READY.");

                // All warnings after SYSTEM READY so they're visible at the bottom
                if (device.HasCompositeFx && device.LeqConfigured)
                {
                    Log("[WARN] CompositeFX keys detected on this device \u2014 blocking LEQ. Click 'Fix Device LEQ' to repair.");
                }
                await CheckClsidHealthAsync();
            }
            catch (Exception ex)
            {
                Log($"[INIT] Failed to read device status: {ex.Message}");
            }
        }

        private AudioDevice? GetSelectedDevice()
        {
            try
            {
                var index = DeviceCombo.SelectedIndex;
                var devices = _devices; // snapshot — prevents race if _devices is replaced mid-read
                if (index < 0 || index >= devices.Count)
                    return null;
                return devices[index];
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        private void UpdateDeviceFormat()
        {
            var device = GetSelectedDevice();
            if (device is null || FormatLayoutText is null || FormatDetailsText is null)
            {
                if (FormatLayoutText != null)
                {
                    FormatLayoutText.Text = "Loading";
                }
                if (FormatDetailsText != null)
                {
                    FormatDetailsText.Text = "...";
                }
                return;
            }

            // Smart Channel Detection: Map channel count to speaker config name
            var layoutName = GetSpeakerConfigName(device.Channels);
            FormatLayoutText.Text = layoutName;
            FormatDetailsText.Text = $"{device.BitDepth}-bit {device.SampleRate}Hz";
        }

        private async Task UpdateIndicatorForSelectedDeviceAsync(IReadOnlyList<AudioDevice>? preloadedDevices = null)
        {
            if (_isUpdatingIndicator) return;
            _isUpdatingIndicator = true;

            try
            {
                var device = GetSelectedDevice();
                if (device is null || string.IsNullOrWhiteSpace(device.Guid))
                {
                    UpdateLeqIndicator(null);
                    UpdateLeqToggleAvailability(null);
                    UpdateReleaseVisuals(4);
                    return;
                }

                try
                {
                    // Use preloaded data when available (e.g. during refresh where we already
                    // have fresh data), otherwise fetch fresh from the service
                    AudioDevice? freshDevice;
                    if (preloadedDevices != null)
                    {
                        freshDevice = preloadedDevices.FirstOrDefault(d => d.Guid == device.Guid);
                    }
                    else
                    {
                        var devices = await _audioService.GetDevicesAsync();
                        freshDevice = devices.FirstOrDefault(d => d.Guid == device.Guid);
                    }

                    if (freshDevice == null)
                    {
                        UpdateLeqIndicator(null);
                        UpdateLeqToggleAvailability(null);
                        UpdateReleaseVisuals(4);
                        return;
                    }

                    // Update the cached devices list with fresh data
                    var deviceIndex = _devices.FindIndex(d => d.Guid == device.Guid);
                    if (deviceIndex >= 0)
                    {
                        _devices[deviceIndex] = freshDevice;
                    }

                    // NOW update toggle availability based on FRESH data
                    UpdateLeqToggleAvailability(freshDevice);

                    // Use preloaded LoudnessEnabled when available to avoid yet another
                    // PowerShell enumeration (GetLeqStateAsync calls GetDevicesAsync internally)
                    bool? state;
                    if (preloadedDevices != null)
                        state = freshDevice.LoudnessEnabled;
                    else
                        state = await _audioService.GetLeqStateAsync(freshDevice.Guid);
                    UpdateLeqIndicator(state);
                    SetReleaseSliderValue(freshDevice.ReleaseTime);
                }
                catch (Exception ex)
                {
                    // Handle device unplug or other device access errors gracefully
                    Log($"Device access error in UpdateIndicatorForSelectedDeviceAsync: {ex.Message}");
                    UpdateLeqIndicator(null); // Reset to unknown state
                    SetReleaseSliderValue(4); // Reset to default

                    // Show user-friendly error if device was unplugged
                    if (ex.Message.Contains("device") || ex.Message.Contains("not found") || ex.Message.Contains("disconnected"))
                    {
                        // Don't spam the user with repeated error messages, just log it
                        // The UI will show "LEQ: N/A" state which is appropriate
                    }
                }
            }
            finally
            {
                _isUpdatingIndicator = false;
            }
        }

        private async Task UpdateEapoStatusAsync()
        {
            if (_isUpdatingEapo) return;
            _isUpdatingEapo = true;

            try
            {
                var device = GetSelectedDevice();
                if (device is null || string.IsNullOrWhiteSpace(device.Guid))
                {
                    UpdateEapoBadge(null);
                    return;
                }

                try
                {
                    var status = await _audioService.GetEapoStatusAsync(device.Guid);
                    UpdateEapoBadge(status);
                }
                catch (Exception)
                {
                    #if DEBUG
                    Debug.WriteLine("UpdateEapoStatusAsync error occurred");
                    #endif
                    UpdateEapoBadge(null);
                }
            }
            finally
            {
                _isUpdatingEapo = false;
            }
        }

        private void SetReleaseSliderValue(int value)
        {
            if (ReleaseSlider == null)
            {
                return;
            }

            _isSliderUpdating = true;
            ReleaseSlider.Value = value;
            UpdateReleaseVisuals(value);
            _isSliderUpdating = false;
            _lastLoggedReleaseValue = value;
        }

        private void UpdateReleaseVisuals(int value)
        {
            var (descriptor, accentColor) = GetReleaseSpectrum(value);
            Brush accentBrush = new SolidColorBrush(accentColor);
            accentBrush.Freeze();

            if (ReleaseSlider != null)
            {
                ReleaseSlider.Foreground = accentBrush;
                ReleaseSlider.BorderBrush = accentBrush;
            }

            if (DigitalReadoutNumber != null)
            {
                // Check if LEQ is OFF
                bool isLeqOff = LeqPowerButton != null && !LeqPowerButton.IsChecked.GetValueOrDefault(false);

                if (isLeqOff)
                {
                    // LEQ is OFF: show "-" and dim the readout
                    DigitalReadoutNumber.Text = "-";
                    DigitalReadoutNumber.Opacity = 0.3;
                }
                else
                {
                    // LEQ is ON: show the actual value
                    DigitalReadoutNumber.Text = value.ToString();
                    DigitalReadoutNumber.Opacity = 1.0;
                }
            }

            if (DigitalReadoutLabel != null)
            {
                bool isLeqOff = LeqPowerButton != null && !LeqPowerButton.IsChecked.GetValueOrDefault(false);

                if (isLeqOff)
                {
                    // LEQ is OFF: dim the label
                    DigitalReadoutLabel.Opacity = 0.3;
                    DigitalReadoutLabel.Text = descriptor; // Keep the descriptor but dimmed
                }
                else
                {
                    // LEQ is ON: show normally
                    DigitalReadoutLabel.Text = descriptor;
                    DigitalReadoutLabel.Opacity = 1.0;
                }
            }
        }

        private static (string descriptor, Color accentColor) GetReleaseSpectrum(int value)
        {
            return value switch
            {
                2 => ("INSTANT", Color.FromRgb(0xFF, 0x45, 0x3A)),
                3 => ("QUICK", Color.FromRgb(0xFF, 0xD6, 0x0A)),
                4 => ("NORMAL", Color.FromRgb(0x32, 0xD7, 0x4B)),
                5 => ("SLOW", Color.FromRgb(0x64, 0xD2, 0xFF)),
                6 => ("SLOWER", Color.FromRgb(0x0A, 0x84, 0xFF)),
                7 => ("SLOWEST", Color.FromRgb(0xBF, 0x5A, 0xF2)),
                _ => ("NORMAL", Color.FromRgb(0x32, 0xD7, 0x4B))
            };
        }

        private void ReleaseSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_isInitialized)
            {
                return;
            }
            if (_isSliderUpdating)
            {
                return;
            }

            var rounded = (int)Math.Round(e.NewValue);
            UpdateReleaseVisuals(rounded);

            if (_lastLoggedReleaseValue != rounded)
            {
                var descriptor = GetReleaseSpectrum(rounded).descriptor;
                Log($"Release Time set to {rounded} ({descriptor})");
                _lastLoggedReleaseValue = rounded;
            }

            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private async void DebounceTimer_Tick(object? sender, EventArgs e)
        {
            _debounceTimer.Stop();
            var device = GetSelectedDevice();
            if (device is null || string.IsNullOrWhiteSpace(device.Guid))
            {
                return;
            }

            var value = ReleaseSlider is null ? 4 : ReleaseSlider.Value;
            // Suspend COM callbacks — registry writes to FxProperties can trigger device change events
            _deviceChangeNotifier?.Suspend();
            try
            {
                await _audioService.SetReleaseTimeAsync(device.Guid, value);
                _devices = _devices.Select(d => d.Guid == device.Guid ? new AudioDevice
                {
                    Guid = d.Guid,
                    Name = d.Name,
                    InterfaceName = d.InterfaceName,
                    SupportsEnhancement = d.SupportsEnhancement,
                    HasLfxGfx = d.HasLfxGfx,
                    HasCompositeFx = d.HasCompositeFx,
                    LeqConfigured = d.LeqConfigured,
                    LoudnessEnabled = d.LoudnessEnabled,
                    ReleaseTime = (int)Math.Round(value),
                    EapoStatus = d.EapoStatus,
                    EapoChildBroken = d.EapoChildBroken,
                    Channels = d.Channels,
                    BitDepth = d.BitDepth,
                    SampleRate = d.SampleRate,
                    HasReleaseTimeKey = d.HasReleaseTimeKey
                } : d).ToList();
            }
            catch (Exception ex)
            {
                var message = ex.Message;
                if (message.Contains("registry access", StringComparison.OrdinalIgnoreCase) ||
                    message.Contains("not allowed", StringComparison.OrdinalIgnoreCase) ||
                    message.Contains("UnauthorizedAccess", StringComparison.OrdinalIgnoreCase))
                {
                    Log("ERROR: Admin rights required to change Release Time. Run as Administrator.");
                }
                else
                {
                    Log($"ERROR: Failed to set Release Time - {message}");
                }
                #if DEBUG
                Debug.WriteLine($"DebounceTimer_Tick error: {ex}");
                #endif
            }
            finally
            {
                ResumeDeviceNotifierAndRefresh();
            }
        }

        private void InitializeTrayIcon()
        {
            _trayIcon = new NotifyIcon();

            try
            {
                // Load icon from embedded resource
                using var iconStream = Application.GetResourceStream(new Uri("pack://application:,,,/icon.ico"))?.Stream;
                if (iconStream != null)
                {
                    _trayIcon.Icon = new Icon(iconStream);
                }
                else
                {
                    throw new InvalidOperationException("Icon resource not found.");
                }
            }
            catch (Exception)
            {
#if DEBUG
                Debug.WriteLine("Icon failed to load");
#endif
                _trayIcon.Icon = SystemIcons.Application;
            }

            _runAtStartupEnabled = IsRunAtStartupEnabled();
            _trayContextMenu = BuildTrayContextMenu();
            _trayIcon.ContextMenuStrip = _trayContextMenu;
            _trayIcon.Visible = true;
            UpdateTrayTooltip();

            _trayContextMenu.Opening += (s, e) =>
            {
                Dispatcher.InvokeAsync(async () =>
                {
                    try
                    {
                        await UpdateTrayMenuState();
                        UpdateDeviceNameMenuText();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[MainWindow] Tray context menu update error: {ex}");
                    }
                });
            };

            _trayIcon.DoubleClick += (s, args) =>
            {
                Dispatcher.Invoke(() =>
                {
                    Show();
                    WindowState = WindowState.Normal;
                    Activate();
                });
            };
        }

        private ContextMenuStrip BuildTrayContextMenu()
        {
            var menu = new ContextMenuStrip
            {
                Renderer = new DarkTrayMenuRenderer()
            };

            var header = new ToolStripMenuItem("LEQ Control Panel")
            {
                Enabled = false,
                Font = new Font(System.Windows.Forms.Control.DefaultFont, System.Drawing.FontStyle.Bold)
            };
            menu.Items.Add(header);

            menu.Items.Add(new ToolStripSeparator());

            _restartAudioServiceItem = new ToolStripMenuItem("Restart Audio Service");
            _restartAudioServiceItem.Click += (_, _) => Dispatcher.InvokeAsync(RestartAudioServiceMenuAsync);
            menu.Items.Add(_restartAudioServiceItem);

            var windowsSoundItem = new ToolStripMenuItem("Win Sound Control Panel");
            windowsSoundItem.Click += (_, _) =>
            {
                try
                {
                    SoundPanelHelper.OpenSoundPanel();
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                        StyledMessageBox.SafeShowError($"Unable to open Windows Sound settings:\n\n{ex.Message}", "Sound Settings"));
                }
            };
            menu.Items.Add(windowsSoundItem);

            _runAtStartupItem = new ToolStripMenuItem("Run at Startup")
            {
                CheckOnClick = true,
                Checked = _runAtStartupEnabled
            };
            _runAtStartupItem.Click += (_, _) => Dispatcher.Invoke(RunAtStartupItem_Click);
            menu.Items.Add(_runAtStartupItem);

            _alwaysOnTopTrayItem = new ToolStripMenuItem("Always on Top")
            {
                CheckOnClick = true,
                Checked = _alwaysOnTop
            };
            _alwaysOnTopTrayItem.Click += (_, _) => {
                _alwaysOnTop = _alwaysOnTopTrayItem.Checked;

                Dispatcher.Invoke(() =>
                {
                    Topmost = _alwaysOnTop;

                    // Update the XAML menu item to stay in sync
                    if (AppSettingsMenu != null)
                    {
                        foreach (var item in AppSettingsMenu.Items)
                        {
                            if (item is MenuItem menuItem && menuItem.Header.ToString() == "Always on Top")
                            {
                                menuItem.IsChecked = _alwaysOnTop;
                                break;
                            }
                        }
                    }
                });

                try
                {
                    using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                    key?.SetValue("AlwaysOnTop", _alwaysOnTop);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[MainWindow] Failed to save AlwaysOnTop setting to registry: {ex.Message}");
                    // Non-critical - setting will work for this session but won't persist
                }
            };
            menu.Items.Add(_alwaysOnTopTrayItem);

            _desktopShortcutTrayItem = new ToolStripMenuItem("Desktop Shortcut")
            {
                CheckOnClick = true,
                Checked = DesktopShortcutExists()
            };
            _desktopShortcutTrayItem.Click += (_, _) =>
            {
                SetDesktopShortcut(_desktopShortcutTrayItem.Checked);
                Dispatcher.Invoke(SyncDesktopShortcutMenuState);
            };
            menu.Items.Add(_desktopShortcutTrayItem);

            _toggleLeqMenuItem = new ToolStripMenuItem("Loudness EQ")
            {
                CheckOnClick = false
            };
            _toggleLeqMenuItem.Click += (_, _) =>
            {
                _toggleLeqMenuItem.Enabled = false;
                Dispatcher.InvokeAsync(async () =>
                {
                    try
                    {
                        await ToggleLeqAsync();
                    }
                    finally
                    {
                        _toggleLeqMenuItem.Enabled = true;
                    }
                });
            };
            menu.Items.Add(_toggleLeqMenuItem);

            // Device friendly name (bold, non-clickable)
            _deviceFriendlyNameMenuItem = new ToolStripMenuItem("No device selected")
            {
                Enabled = false,
                ForeColor = System.Drawing.Color.FromArgb(220, 220, 220),
                Font = new Font(System.Windows.Forms.Control.DefaultFont.FontFamily, 9.5f, System.Drawing.FontStyle.Bold),
                Padding = new Padding(0, 0, 0, -4),
            };
            menu.Items.Add(_deviceFriendlyNameMenuItem);

            // Device descriptor (smaller italic, non-clickable)
            _deviceDescriptorMenuItem = new ToolStripMenuItem("")
            {
                Enabled = false,
                ForeColor = System.Drawing.Color.FromArgb(128, 128, 128),
                Font = new Font(System.Windows.Forms.Control.DefaultFont.FontFamily, 8.0f, System.Drawing.FontStyle.Italic),
                Padding = new Padding(0, -4, 0, 0),
            };
            menu.Items.Add(_deviceDescriptorMenuItem);

            // Update text immediately based on current state
            UpdateDeviceNameMenuText();
            // Note: Async tray menu updates happen when menu opens

            menu.Items.Add(new ToolStripSeparator());

            var showInterfaceItem = new ToolStripMenuItem("Show Interface");
            showInterfaceItem.Click += (_, _) =>
            {
                Show();
                WindowState = WindowState.Normal;
                Activate();
            };
            menu.Items.Add(showInterfaceItem);

            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (_, _) => ExitApplication();
            menu.Items.Add(exitItem);

            return menu;
        }

        private void RunAtStartupItem_Click()
        {
            var enable = _runAtStartupItem.Checked;
            try
            {
                SetRunAtStartup(enable);
            }
            catch (Exception ex)
            {
                StyledMessageBox.SafeShowError($"Unable to update startup setting:\n\n{ex.Message}", "Run at Startup");
                _runAtStartupItem.Checked = !enable;
            }
        }

        private void SetRunAtStartup(bool enable)
        {
            var path = GetExecutablePath();
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new InvalidOperationException("Executable path could not be determined.");
            }

            using var key = Registry.CurrentUser.CreateSubKey(RunRegistryPath, true);
            if (key is null)
            {
                throw new InvalidOperationException("Unable to open Run registry key.");
            }

            if (enable)
            {
                key.SetValue(RunValueName, $"\"{path}\"");
            }
            else
            {
                key.DeleteValue(RunValueName, false);
            }

            _runAtStartupEnabled = enable;
        }

        private bool IsRunAtStartupEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunRegistryPath, false);
                var value = key?.GetValue(RunValueName)?.ToString();
                var path = GetExecutablePath();
                if (string.IsNullOrWhiteSpace(path))
                {
                    return false;
                }

                return string.Equals(value, $"\"{path}\"", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(value, path, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[MainWindow] Failed to check startup registry: {ex.Message}");
                return false;
            }
        }

        private string GetExecutablePath()
        {
            // In single-file publish, Environment.ProcessPath should work
            // But as a fallback, try to get the current process executable path
            try
            {
                if (Environment.ProcessPath is { } processPath)
                    return processPath;

                using var currentProcess = System.Diagnostics.Process.GetCurrentProcess();
                return currentProcess.MainModule?.FileName ??
                       System.AppContext.BaseDirectory ??
                       string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[MainWindow] GetExecutablePath failed: {ex.Message}");
                return string.Empty;
            }
        }

        private string GetDesktopShortcutPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "LEQ Control Panel.lnk");
        }

        private bool DesktopShortcutExists()
        {
            return File.Exists(GetDesktopShortcutPath());
        }

        private void SetDesktopShortcut(bool create)
        {
            var shortcutPath = GetDesktopShortcutPath();
            if (create)
            {
                CreateOrUpdateDesktopShortcut(shortcutPath);
            }
            else
            {
                if (File.Exists(shortcutPath))
                    File.Delete(shortcutPath);
            }
        }

        private void CreateOrUpdateDesktopShortcut(string shortcutPath)
        {
            var exePath = GetExecutablePath();
            if (string.IsNullOrWhiteSpace(exePath)) return;

            var shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null) return;
            var shellObj = Activator.CreateInstance(shellType);
            if (shellObj == null) return;
            dynamic shell = shellObj;
            try
            {
                dynamic shortcut = shell.CreateShortcut(shortcutPath);
                try
                {
                    shortcut.TargetPath = exePath;
                    shortcut.WorkingDirectory = Path.GetDirectoryName(exePath) ?? "";
                    shortcut.Description = "LEQ Control Panel";
                    shortcut.Save();
                }
                finally
                {
                    Marshal.FinalReleaseComObject(shortcut);
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(shell);
            }
        }

        private void SyncDesktopShortcutMenuState()
        {
            var exists = DesktopShortcutExists();
            if (_desktopShortcutTrayItem != null)
                _desktopShortcutTrayItem.Checked = exists;
            if (MenuDesktopShortcut != null)
                MenuDesktopShortcut.IsChecked = exists;
        }

        private void Setting_DesktopShortcut_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                try
                {
                    SetDesktopShortcut(menuItem.IsChecked);
                    SyncDesktopShortcutMenuState();
                }
                catch (Exception ex)
                {
                    StyledMessageBox.SafeShowError($"Unable to update desktop shortcut:\n\n{ex.Message}", "Desktop Shortcut");
                    menuItem.IsChecked = !menuItem.IsChecked;
                    SyncDesktopShortcutMenuState();
                }
            }
        }

        private void UpdateStalePathsOnStartup()
        {
            var currentPath = GetExecutablePath();
            if (string.IsNullOrWhiteSpace(currentPath)) return;

            // Update stale desktop shortcut
            try
            {
                var shortcutPath = GetDesktopShortcutPath();
                if (File.Exists(shortcutPath))
                {
                    CreateOrUpdateDesktopShortcut(shortcutPath);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[MainWindow] Failed to update desktop shortcut path: {ex.Message}");
            }

            // Update stale startup registry entry
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunRegistryPath, false);
                var value = key?.GetValue(RunValueName)?.ToString();
                if (!string.IsNullOrEmpty(value))
                {
                    var expectedValue = $"\"{currentPath}\"";
                    if (!string.Equals(value, expectedValue, StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(value, currentPath, StringComparison.OrdinalIgnoreCase))
                    {
                        SetRunAtStartup(true);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[MainWindow] Failed to update startup registry path: {ex.Message}");
            }
        }

        private async Task RestartAudioServiceMenuAsync()
        {
            if (_restartAudioServiceItem == null)
            {
                return;
            }

            _restartAudioServiceItem.Enabled = false;
            try
            {
                var (restarted, reason) = await RestartAudioServiceWithSuspendAsync();
                if (restarted)
                    StyledMessageBox.ShowInfo("Audio service restarted successfully.", "Audio Service");
                else
                    StyledMessageBox.ShowWarning($"Unable to restart the Windows audio service.\n\n{reason}", "Audio Service");
            }
            catch (Exception ex)
            {
                StyledMessageBox.SafeShowError($"Audio service restart failed:\n\n{ex.Message}", "Audio Service");
            }
            finally
            {
                _restartAudioServiceItem.Enabled = true;
                UpdateEapoInstalledState();
            }
        }

        private async Task ToggleLeqAsync()
        {
            var device = GetSelectedDevice();
            if (device is null || string.IsNullOrWhiteSpace(device.Guid))
            {
                return;
            }

            try
            {
                // Toggle uses direct .NET Registry API; requires admin privileges.
                // Always refresh state from the service afterward so the UI reflects reality.
                _ = await _audioService.ToggleLeqAsync(device.Guid);
                var state = await _audioService.GetLeqStateAsync(device.Guid);
                UpdateLeqIndicator(state);
                await UpdateEapoStatusAsync();
                Log(state switch
                {
                    true => "LEQ enabled.",
                    false => "LEQ disabled.",
                    null => "LEQ state unknown (toggle may have been cancelled)."
                });

                // Update tray menu text to reflect new state
                await UpdateTrayMenuState();
            }
            catch (Exception ex)
            {
                // Handle device unplug or other device access errors gracefully
                Log($"Device access error in ToggleLeqAsync: {ex.Message}");
                UpdateLeqIndicator(null); // Reset to unknown state
                await UpdateEapoStatusAsync(); // This will handle its own errors

                // Show user-friendly error if device was unplugged
                if (ex.Message.Contains("device") || ex.Message.Contains("not found") || ex.Message.Contains("disconnected"))
                {
                    StyledMessageBox.SafeShowWarning(
                        "The selected audio device is no longer available.\n\n" +
                        "This can happen if the device was unplugged or disconnected.\n\n" +
                        "Please select a different device from the dropdown.",
                        "Device Unavailable");
                }
                else
                {
                    StyledMessageBox.SafeShowError(
                        $"Failed to toggle LEQ:\n\n{ex.Message}\n\n" +
                        "Check the console for more details.",
                        "Toggle Failed");
                }
            }
        }

        internal void DisposeTrayIcon()
        {
            try
            {
                if (_trayIcon != null)
                {
                    _trayIcon.Visible = false;
                    _trayIcon.Dispose();
                }
            }
            catch
            {
                // Best-effort — app state may be corrupt during crash
            }
        }

        internal void ExitApplication()
        {
            _isExiting = true;
            try { _debounceTimer.Stop(); } catch { }
            try { _deviceChangeDebounceTimer.Stop(); } catch { }

            if (_deviceChangeNotifier != null)
            {
                try { _deviceChangeNotifier.DeviceChanged -= OnExternalDeviceChanged; } catch { }
                try { _deviceChangeNotifier.Dispose(); } catch { }
                _deviceChangeNotifier = null;
            }

            try { DisposeTrayIcon(); } catch { }
            Application.Current.Shutdown();
        }

        private async Task UpdateAllUIState()
        {
            // Main window updates
            await UpdateIndicatorForSelectedDeviceAsync();
            await UpdateEapoStatusAsync();
            UpdateDeviceFormat();

            // Tray menu updates
            await UpdateTrayMenuState();
        }

        // Event handlers for overlay animations to prevent accumulation
        private void OnFadeOutLoadingOverlayCompleted(object? sender, EventArgs e)
        {
            if (LeqLoadingOverlay != null)
            {
                LeqLoadingOverlay.IsHitTestVisible = false;
            }
        }

        private void ShowLeqLoadingOverlay(string text)
        {
            if (LoadingOverlayText != null)
                LoadingOverlayText.Text = text;
            if (LeqLoadingOverlay != null)
                LeqLoadingOverlay.IsHitTestVisible = true;
            var fadeIn = (Storyboard)FindResource("FadeInLoadingOverlay");
            fadeIn.Begin();
        }

        private void HideLeqLoadingOverlay()
        {
            try
            {
                var fadeOut = (Storyboard)FindResource("FadeOutLoadingOverlay");
                fadeOut.Completed -= OnFadeOutLoadingOverlayCompleted;
                fadeOut.Completed += OnFadeOutLoadingOverlayCompleted;
                fadeOut.Begin();
            }
            catch
            {
                if (LeqLoadingOverlay != null)
                {
                    LeqLoadingOverlay.Opacity = 0;
                    LeqLoadingOverlay.IsHitTestVisible = false;
                }
            }
        }

        /// <summary>
        /// Animates an overlay's opacity to show or hide it with a smooth fade.
        /// </summary>
        private void FadeOverlay(UIElement overlay, bool show, double duration = 0.2)
        {
            if (overlay == null) return;

            double toValue = show ? 1.0 : 0.0;

            // Skip if already at target
            if (Math.Abs(overlay.Opacity - toValue) < 0.01)
            {
                overlay.IsHitTestVisible = show;
                return;
            }

            var animation = new System.Windows.Media.Animation.DoubleAnimation
            {
                To = toValue,
                Duration = TimeSpan.FromSeconds(duration),
                EasingFunction = new System.Windows.Media.Animation.SineEase
                {
                    EasingMode = System.Windows.Media.Animation.EasingMode.EaseInOut
                }
            };

            if (show)
            {
                overlay.IsHitTestVisible = true;
            }
            else
            {
                animation.Completed += (s, e) =>
                {
                    overlay.IsHitTestVisible = false;
                };
            }

            overlay.BeginAnimation(UIElement.OpacityProperty, animation);
        }

        private Storyboard? _deviceLoadingBarStoryboard;

        private void StartDeviceLoadingBar()
        {
            // Start the shimmer animation on the fill element
            _deviceLoadingBarStoryboard = new Storyboard { RepeatBehavior = RepeatBehavior.Forever };
            var shimmer = new ThicknessAnimation
            {
                From = new Thickness(-80, 0, 0, 0),
                To = new Thickness(220, 0, 0, 0),
                Duration = TimeSpan.FromSeconds(1.5)
            };
            Storyboard.SetTarget(shimmer, DeviceLoadingBarFill);
            Storyboard.SetTargetProperty(shimmer, new PropertyPath(MarginProperty));
            _deviceLoadingBarStoryboard.Children.Add(shimmer);
            _deviceLoadingBarStoryboard.Begin();

            // Fade in the bar
            var fadeIn = (Storyboard)FindResource("FadeInDeviceLoadingBar");
            fadeIn.Begin();
        }

        private void StopDeviceLoadingBar()
        {
            try
            {
                // Fade out the bar
                var fadeOut = (Storyboard)FindResource("FadeOutDeviceLoadingBar");
                fadeOut.Completed -= OnFadeOutDeviceLoadingBarCompleted;
                fadeOut.Completed += OnFadeOutDeviceLoadingBarCompleted;
                fadeOut.Begin();
            }
            catch
            {
                // Fallback: directly stop the animation
                _deviceLoadingBarStoryboard?.Stop();
                _deviceLoadingBarStoryboard = null;
            }
        }

        private void OnFadeOutDeviceLoadingBarCompleted(object? sender, EventArgs e)
        {
            _deviceLoadingBarStoryboard?.Stop();
            _deviceLoadingBarStoryboard = null;
        }

        private async Task UpdateTrayMenuState()
        {
            // Update device name
            UpdateDeviceNameMenuText();

            // Update LEQ toggle state
            await UpdateToggleLeqMenuText();
        }

        private async Task UpdateToggleLeqMenuText()
        {
            if (_toggleLeqMenuItem == null) return;

            var selectedDevice = DeviceCombo.SelectedItem as AudioDevice;
            if (selectedDevice == null)
            {
                _toggleLeqMenuItem.Text = "Loudness EQ: N/A";
                _toggleLeqMenuItem.Enabled = false;
                return;
            }

            // Check if LEQ is installed/configured on this device
            if (!selectedDevice.LeqConfigured)
            {
                // LEQ not installed - disable and show not installed message
                _toggleLeqMenuItem.Text = "Loudness EQ (Not Installed)";
                _toggleLeqMenuItem.Enabled = false;
                _toggleLeqMenuItem.ForeColor = System.Drawing.Color.FromArgb(136, 136, 136); // Grey out
                return;
            }

            // LEQ is installed - enable and show current state
            _toggleLeqMenuItem.Enabled = true;
            _toggleLeqMenuItem.ForeColor = System.Drawing.Color.White; // Normal color

            try
            {
                // Get fresh device info to check actual current state
                var leqState = await _audioService.GetLeqStateAsync(selectedDevice.Guid);

                if (leqState.HasValue && leqState.Value)
                {
                    _toggleLeqMenuItem.Text = "Loudness EQ: ON";
                }
                else
                {
                    _toggleLeqMenuItem.Text = "Loudness EQ: OFF";
                }
            }
            catch
            {
                // If we can't get state, default to OFF
                _toggleLeqMenuItem.Text = "Loudness EQ: OFF";
            }
        }

        private void UpdateDeviceNameMenuText()
        {
            if (_deviceFriendlyNameMenuItem == null) return;

            var device = GetSelectedDevice();
            if (device != null && !string.IsNullOrWhiteSpace(device.Name))
            {
                _deviceFriendlyNameMenuItem.Text = $"    {device.Name}";
                _deviceDescriptorMenuItem.Text = $"    {device.InterfaceName}";
                _deviceDescriptorMenuItem.Visible = !string.IsNullOrWhiteSpace(device.InterfaceName);
            }
            else
            {
                _deviceFriendlyNameMenuItem.Text = "    No device selected";
                _deviceDescriptorMenuItem.Visible = false;
            }
        }

        private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ButtonState == System.Windows.Input.MouseButtonState.Pressed)
            {
                // Only allow window dragging if click is in the top 30 pixels (header row)
                var mousePos = e.GetPosition(this);
                if (mousePos.Y <= 30)
                {
                    DragMove();
                }
            }
        }

        /// <summary>
        /// Resumes device change notifications and triggers a device list refresh
        /// to catch any changes that were missed while the notifier was suspended.
        /// </summary>
        private void ResumeDeviceNotifierAndRefresh()
        {
            _deviceChangeNotifier?.Resume();
            _ = RefreshDeviceListAsync();
        }

        protected override void OnStateChanged(EventArgs e)
        {
            if (WindowState == WindowState.Minimized)
            {
                SuppressGlowAnimations();
                Hide();
            }
            else if (WindowState == WindowState.Normal)
            {
                RestoreGlowAnimations();
            }

            base.OnStateChanged(e);
        }

        /// <summary>
        /// Temporarily disables the Install button to stop trigger-driven glow storyboards,
        /// preventing CPU usage from software-rendered animations while the window is hidden.
        /// </summary>
        private void SuppressGlowAnimations()
        {
            if (_glowAnimationsSuppressed) return;
            _glowAnimationsSuppressed = true;

            if (InstallLeqButton != null)
                InstallLeqButton.IsEnabled = false;
        }

        /// <summary>
        /// Re-evaluates the Install button state after the window becomes visible,
        /// which re-triggers glow storyboards if the button should be enabled.
        /// </summary>
        private void RestoreGlowAnimations()
        {
            if (!_glowAnimationsSuppressed) return;
            _glowAnimationsSuppressed = false;

            var device = GetSelectedDevice();
            if (device != null)
                UpdateInstallButtonState(device.LeqConfigured, device.HasLfxGfx, device.HasCompositeFx, _clsidsBroken);
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (_isExiting)
            {
                base.OnClosing(e);
                return;
            }

            if (_closeBehaviorRemembered)
            {
                // Behavior already chosen — apply directly
                e.Cancel = true;
                ApplyCloseBehavior(_closeBehavior);
            }
            else
            {
                // Ask the user what they want to do
                e.Cancel = true;
                ShowCloseBehaviorDialog();
            }
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void LoadApplicationSettings()
        {
            try
            {
                // Load settings from registry
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(SettingsRegistryPath);
                if (key != null)
                {
                    _startMinimized = Convert.ToBoolean(key.GetValue("StartMinimized", false));
                    var closeBehaviorStr = key.GetValue("CloseBehavior", "")?.ToString() ?? "";
                    if (Enum.TryParse<CloseBehavior>(closeBehaviorStr, out var parsedBehavior))
                    {
                        _closeBehavior = parsedBehavior;
                        _closeBehaviorRemembered = true;
                    }
                    else
                    {
                        // Migrate legacy CloseToTray boolean
                        var legacyCloseToTray = Convert.ToBoolean(key.GetValue("CloseToTray", false));
                        if (legacyCloseToTray)
                        {
                            _closeBehavior = CloseBehavior.MinimizeToTray;
                            _closeBehaviorRemembered = true;
                        }
                    }
                    _skipEapoWarning = Convert.ToBoolean(key.GetValue("SkipEapoWarning", false));
                    _skipResetDeviceWarning = Convert.ToBoolean(key.GetValue("SkipResetDeviceWarning", false));
                    _alwaysOnTop = Convert.ToBoolean(key.GetValue("AlwaysOnTop", false));
                    _lastSelectedDeviceGuid = key.GetValue("LastSelectedDeviceGuid", "") as string ?? "";
                }

                // Apply always on top setting
                Topmost = _alwaysOnTop;

                // Update menu items to reflect current settings
                if (AppSettingsMenu != null)
                {
                    foreach (var item in AppSettingsMenu.Items)
                    {
                        if (item is MenuItem menuItem)
                        {
                            switch (menuItem.Header.ToString())
                            {
                                case "Run at Startup":
                                    menuItem.IsChecked = _runAtStartupEnabled;
                                    break;
                                case "Start Minimized":
                                    menuItem.IsChecked = _startMinimized;
                                    break;
                                // Close behavior submenu items are handled by UpdateCloseBehaviorMenuItems()
                                case "Always on Top":
                                    menuItem.IsChecked = _alwaysOnTop;
                                    break;
                                case "Desktop Shortcut":
                                    menuItem.IsChecked = DesktopShortcutExists();
                                    break;
                            }
                        }
                    }
                }

                // Update close behavior submenu items
                UpdateCloseBehaviorMenuItems();
            }
            catch (Exception ex)
            {
                Log($"Failed to load application settings: {ex.Message}");
            }
        }

        private void ConsoleToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (ConsoleList.Visibility == Visibility.Visible)
            {
                // Collapse console
                ConsoleList.Visibility = Visibility.Collapsed;
                ConsoleToggleIcon.Text = "\u25B2";
                Height -= 120; // Reduce window height by approximate console height
            }
            else
            {
                // Expand console
                ConsoleList.Visibility = Visibility.Visible;
                ConsoleToggleIcon.Text = "\u25BC";
                Height += 120; // Increase window height by approximate console height
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            if (_closeBehaviorRemembered)
            {
                ApplyCloseBehavior(_closeBehavior);
            }
            else
            {
                ShowCloseBehaviorDialog();
            }
        }

        private void ShowCloseBehaviorDialog()
        {
            var dialog = new ConfirmationDialog();
            dialog.Owner = this;
            var result = dialog.ShowDialog();

            if (result == true)
            {
                _closeBehavior = dialog.SelectedBehavior;

                if (dialog.RememberChoice)
                {
                    _closeBehaviorRemembered = true;
                    SaveCloseBehaviorSetting();
                    UpdateCloseBehaviorMenuItems();
                }

                ApplyCloseBehavior(_closeBehavior);
            }
        }

        private void ApplyCloseBehavior(CloseBehavior behavior)
        {
            switch (behavior)
            {
                case CloseBehavior.MinimizeToTray:
                    Hide();
                    _trayIcon?.ShowBalloonTip(2000, "LEQ Control Panel", "Minimized to tray", ToolTipIcon.Info);
                    break;
                case CloseBehavior.ExitApplication:
                    ExitApplication();
                    break;
            }
        }

        private void SaveCloseBehaviorSetting()
        {
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                key?.SetValue("CloseBehavior", _closeBehavior.ToString());
            }
            catch (Exception ex)
            {
                Log($"Failed to save close behavior setting: {ex.Message}");
            }
        }

        private void AppSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            // Open the context menu
            if (AppSettingsMenu != null)
            {
                AppSettingsMenu.IsOpen = true;
            }
        }

        private void Setting_RunAtStartup_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                _runAtStartupEnabled = menuItem.IsChecked;

                // Sync tray menu
                if (_runAtStartupItem != null)
                {
                    _runAtStartupItem.Checked = _runAtStartupEnabled;
                }

                try
                {
                    // Write the Run key first — it's the source of truth for Windows startup behavior.
                    // Only update the app setting mirror if this succeeds.
                    using var startupKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
                    if (startupKey != null)
                    {
                        if (_runAtStartupEnabled)
                        {
                            var exePath = Environment.ProcessPath ?? System.AppContext.BaseDirectory;
                            startupKey.SetValue(RunValueName, $"\"{exePath}\"");
                        }
                        else
                        {
                            startupKey.DeleteValue(RunValueName, false);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Failed to update startup settings: {ex.Message}");
                }
            }
        }

        private void Setting_StartMin_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                _startMinimized = menuItem.IsChecked;
                try
                {
                    using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                    key?.SetValue("StartMinimized", _startMinimized);
                }
                catch (Exception ex)
                {
                    Log($"Failed to save start minimized setting: {ex.Message}");
                }
            }
        }

        private void Setting_CloseBehavior_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                switch (menuItem.Name)
                {
                    case "MenuMinimizeToTray":
                        _closeBehavior = CloseBehavior.MinimizeToTray;
                        _closeBehaviorRemembered = true;
                        break;
                    case "MenuExitApplication":
                        _closeBehavior = CloseBehavior.ExitApplication;
                        _closeBehaviorRemembered = true;
                        break;
                    case "MenuAskEveryTime":
                        _closeBehaviorRemembered = false;
                        break;
                }

                UpdateCloseBehaviorMenuItems();

                if (_closeBehaviorRemembered)
                {
                    SaveCloseBehaviorSetting();
                }
                else
                {
                    // Clear saved behavior so the dialog shows again
                    try
                    {
                        using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                        key?.DeleteValue("CloseBehavior", false);
                    }
                    catch (Exception ex)
                    {
                        Log($"Failed to clear close behavior setting: {ex.Message}");
                    }
                }
            }
        }

        private void UpdateCloseBehaviorMenuItems()
        {
            if (MenuMinimizeToTray != null)
                MenuMinimizeToTray.IsChecked = _closeBehaviorRemembered && _closeBehavior == CloseBehavior.MinimizeToTray;
            if (MenuExitApplication != null)
                MenuExitApplication.IsChecked = _closeBehaviorRemembered && _closeBehavior == CloseBehavior.ExitApplication;
            if (MenuAskEveryTime != null)
                MenuAskEveryTime.IsChecked = !_closeBehaviorRemembered;
        }

        private void Setting_AlwaysOnTop_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                _alwaysOnTop = menuItem.IsChecked;
                Topmost = _alwaysOnTop;

                // Sync tray menu
                if (_alwaysOnTopTrayItem != null)
                {
                    _alwaysOnTopTrayItem.Checked = _alwaysOnTop;
                }

                try
                {
                    using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(SettingsRegistryPath);
                    key?.SetValue("AlwaysOnTop", _alwaysOnTop);
                }
                catch (Exception ex)
                {
                    Log($"Failed to save always on top setting: {ex.Message}");
                }
            }
        }

        private async Task CheckClsidHealthAsync()
        {
            try
            {
                var broken = await Task.Run(() => _audioService.CheckClsidHealth());
                _clsidsBroken = broken.Length > 0;

                if (_clsidsBroken)
                {
                    Log($"[WARN] CLSID registrations broken: {string.Join(", ", broken)}. Use Settings \u25b8 Fix LEQ Registry to repair.");
                }

                // Refresh LEQ controls to reflect CLSID state
                var device = GetSelectedDevice();
                if (device != null) UpdateLeqToggleAvailability(device);
            }
            catch
            {
                // Don't block startup if check fails
            }
        }

        private async Task FixClsidFromButtonAsync()
        {
            Log("[REGISTRY] Checking CLSID registrations...");

            try
            {
                var (missing, fixed_, failed) = await _audioService.RepairClsidsAsync();

                if (missing.Length == 0)
                {
                    Log("[REGISTRY] All CLSIDs are registered correctly. No fix needed.");
                    _clsidsBroken = false;
                    await CheckClsidHealthAsync();
                    return;
                }

                if (fixed_.Length > 0)
                {
                    Log($"[REGISTRY] Fixed missing/broken CLSIDs: {string.Join(", ", fixed_)}");
                }

                if (failed.Length > 0)
                {
                    Log($"[REGISTRY] \u26a0 Failed to fix CLSIDs: {string.Join(", ", failed)}. Run as administrator.");
                }

                if (fixed_.Length > 0)
                {
                    Log("[REGISTRY] Restarting audio service...");
                    var (success, reason) = await RestartAudioServiceWithSuspendAsync();
                    if (success)
                    {
                        Log("[REGISTRY] Audio service restarted. Fix complete.");
                    }
                    else
                    {
                        Log($"[REGISTRY] \u26a0 Audio service restart failed: {reason}");
                    }

                    // Re-check CLSID health and refresh device list
                    await CheckClsidHealthAsync();
                    await RefreshDeviceListAsync();
                }
            }
            catch (Exception ex)
            {
                Log($"[REGISTRY] \u26a0 Registry repair failed: {ex.Message}");
            }
        }

        private async void FixLeqRegistry_Click(object sender, RoutedEventArgs e)
        {
            await FixClsidFromButtonAsync();
        }

        private async Task FixDeviceCompositeFxAsync(AudioDevice device)
        {
            Log("[FIX] Removing CompositeFX keys from device...");

            // Show loading overlay
            if (LeqLoadingOverlay != null && LoadingOverlayText != null)
            {
                LoadingOverlayText.Text = "Fixing device LEQ registry...";
                FadeOverlay(LeqLoadingOverlay, true);
            }

            InstallLeqButton.IsEnabled = false;

            try
            {
                var success = await _audioService.FixDeviceCompositeFxAsync(device.Guid);

                if (success)
                {
                    Log("[FIX] CompositeFX keys removed successfully.");
                    Log("[FIX] Restarting audio service...");

                    var (restartOk, reason) = await RestartAudioServiceWithSuspendAsync();
                    if (restartOk)
                    {
                        Log("[FIX] Audio service restarted. Fix complete.");
                    }
                    else
                    {
                        Log($"[FIX] \u26a0 Audio service restart failed: {reason}");
                    }

                    // Refresh device list to pick up the cleared CompositeFX state
                    await RefreshDeviceListAsync();
                }
                else
                {
                    Log("[FIX] \u26a0 Failed to remove CompositeFX keys.");
                }
            }
            catch (Exception ex)
            {
                Log($"[FIX] \u26a0 CompositeFX fix failed: {ex.Message}");
            }
            finally
            {
                if (LeqLoadingOverlay != null)
                    FadeOverlay(LeqLoadingOverlay, false);
                InstallLeqButton.IsEnabled = true;
            }
        }

        private async void CheckForUpdates_Click(object sender, RoutedEventArgs e)
        {
            Log("Checking for updates...");

            try
            {
                var result = await UpdateChecker.CheckForUpdateAsync();
                var currentVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.0.0";

                if (result == null)
                {
                    StyledMessageBox.ShowWarning(
                        "Could not check for updates right now.\n\nPlease try again later.",
                        "Update Check Failed");
                    return;
                }

                if (result.Value.UpdateAvailable)
                {
                    Log($"Update available: v{result.Value.NewVersion}");
                    var dialogResult = StyledMessageBox.ShowYesNo(
                        $"A new version is available!\n\n" +
                        $"Current version: v{currentVersion}\n" +
                        $"Latest version: v{result.Value.NewVersion}\n\n" +
                        $"Would you like to download and install it now?",
                        "Update Available");

                    if (dialogResult == MessageBoxResult.Yes)
                    {
                        Log("Downloading update...");
                        bool relaunching = await UpdateService.DownloadAndSwapAsync(
                            result.Value.DownloadUrl, result.Value.NewVersion, this, result.Value.Sha256Hash);

                        if (relaunching)
                        {
                            Log("Update downloaded, relaunching...");
                            ExitApplication();
                        }
                        else
                        {
                            Log("Update cancelled or failed.");
                        }
                    }
                }
                else
                {
                    Log("No updates available - you're running the latest version.");
                    StyledMessageBox.ShowInfo($"\uD83C\uDF89 You're all set!\n\nYou're running the latest version (v{currentVersion}).\n\nNo updates are currently available.", "Up to Date");
                }
            }
            catch (Exception ex)
            {
                Log($"Update check failed: {ex.Message}");
                StyledMessageBox.SafeShowWarning(
                    $"Could not check for updates.\n\n{ex.Message}\n\nPlease check manually at the GitHub releases page.",
                    "Update Check Failed");
            }
        }

        private static string GetSpeakerConfigName(int channels)
        {
            return channels switch
            {
                1 => "Mono",
                2 => "Stereo",
                4 => "Quadraphonic",
                6 => "5.1 Surround",
                8 => "7.1 Surround",
                _ => $"{channels} Channels"
            };
        }

        private static (bool Installed, string? ExePath) DetectArtTuneKit()
        {
            // 1. Bundled mode — same or parent directory
            var localPath = Path.Combine(AppContext.BaseDirectory, "ArtTuneKit.exe");
            if (File.Exists(localPath)) return (true, localPath);

            var parentDir = Directory.GetParent(AppContext.BaseDirectory)?.FullName;
            if (!string.IsNullOrEmpty(parentDir))
            {
                var parentPath = Path.Combine(parentDir, "ArtTuneKit.exe");
                if (File.Exists(parentPath)) return (true, parentPath);
            }

            // 2. Default install location
            var programFiles = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "ArtTuneKit", "ArtTuneKit.exe");
            if (File.Exists(programFiles)) return (true, programFiles);

            // 3. Registry uninstall key
            try
            {
                using var uninstallKey = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                if (uninstallKey != null)
                {
                    foreach (var subKeyName in uninstallKey.GetSubKeyNames())
                    {
                        using var subKey = uninstallKey.OpenSubKey(subKeyName);
                        var displayName = subKey?.GetValue("DisplayName")?.ToString();
                        if (displayName != null &&
                            displayName.Contains("ArtTuneKit", StringComparison.OrdinalIgnoreCase))
                        {
                            var installLocation = subKey?.GetValue("InstallLocation")?.ToString();
                            if (!string.IsNullOrEmpty(installLocation))
                            {
                                var atkPath = Path.Combine(installLocation, "ArtTuneKit.exe");
                                if (File.Exists(atkPath)) return (true, atkPath);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[MainWindow] DetectArtTuneKit registry check failed: {ex.Message}");
            }

            return (false, null);
        }

        private void RefreshArtTuneKitDetection()
        {
            var (installed, exePath) = DetectArtTuneKit();
            _atkInstalled = installed;
            _atkExePath = exePath;
        }

        private static bool IsArtTuneDevice(AudioDevice? device)
        {
            if (device == null) return false;
            return device.Name.StartsWith("Art Tune", StringComparison.OrdinalIgnoreCase);
        }

        private void UpdateArtTuneGating()
        {
            var device = GetSelectedDevice();
            bool gated = IsArtTuneDevice(device) && _atkInstalled;

            // Show/hide ATK overlay on the LEQ control unit
            FadeOverlay(AtkManagedOverlay, gated);

            // Dim the selected device display in the ComboBox face when gated
            DeviceCombo.Opacity = gated ? 0.4 : 1.0;

            // Grey out Art Tune devices in the ComboBox when ATK is installed
            for (int i = 0; i < DeviceCombo.Items.Count; i++)
            {
                var container = DeviceCombo.ItemContainerGenerator.ContainerFromIndex(i) as System.Windows.Controls.ComboBoxItem;
                if (container == null) continue;
                var dev = DeviceCombo.Items[i] as AudioDevice;
                if (dev != null && IsArtTuneDevice(dev) && _atkInstalled)
                {
                    container.Opacity = 0.4;
                    container.ToolTip = "Managed by ArtTuneKit";
                }
                else
                {
                    container.Opacity = 1.0;
                    container.ToolTip = null;
                }
            }

            if (!gated) return;

            // When gated, disable LEQ controls (override their normal state)
            if (_isAdmin)
            {
                InstallLeqButton.IsEnabled = false;
                InstallLeqButton.Opacity = 0.5;
                InstallLeqButton.ToolTip = "This device is managed by ArtTuneKit";

                CleanInstallCheck.IsEnabled = false;
                CleanInstallCheck.Opacity = 0.5;

                ResetDeviceButton.IsEnabled = false;
                ResetDeviceButton.Opacity = 0.5;
                ResetDeviceButton.ToolTip = "This device is managed by ArtTuneKit";
            }

            // Hide the "LEQ Not Detected" overlay — the ATK overlay replaces it
            FadeOverlay(LeqNotDetectedOverlay, false);
        }

        private void OpenArtTuneKit_Click(object sender, RoutedEventArgs e)
        {
            var exePath = _atkExePath;
            if (string.IsNullOrEmpty(exePath))
            {
                // Re-detect in case path changed
                RefreshArtTuneKitDetection();
                exePath = _atkExePath;
            }

            if (string.IsNullOrEmpty(exePath))
            {
                StyledMessageBox.ShowError(
                    "Couldn't find ArtTuneKit executable.",
                    "Launch Failed");
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo(exePath) { UseShellExecute = true })?.Dispose();

                // Shut down SCP cleanly after launching ATK
                Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                StyledMessageBox.SafeShowError(
                    $"Couldn't launch ArtTuneKit: {ex.Message}",
                    "Launch Failed");
            }
        }
        }
    }