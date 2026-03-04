// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Application = System.Windows.Application;
using LEQControlPanel;
using LEQControlPanel.Dialogs;
using LEQControlPanel.Services;
using LEQControlPanel.Windows;

namespace LEQControlPanel;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    private const string MutexName = "Global\\LEQControlPanel-SingleInstance";
    private static Mutex? _instanceMutex;
    private static bool _mutexOwned;

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    private const int SW_RESTORE = 9;

    protected override async void OnStartup(StartupEventArgs e)
    {
        // Add global exception handler for startup issues
        AppDomain.CurrentDomain.UnhandledException += (sender, args) =>
        {
            (Application.Current.MainWindow as MainWindow)?.DisposeTrayIcon();
            try { if (_instanceMutex != null && _mutexOwned) _instanceMutex.ReleaseMutex(); } catch { }
            var exception = args.ExceptionObject as Exception;
            Debug.WriteLine($"[LEQControlPanel] Critical startup error: {exception}");
            try
            {
                System.Windows.MessageBox.Show(
                    "An unexpected error occurred and LEQ Control Panel needs to close.\n\n" +
                    "Please restart the application. If this continues, submit a diagnostic report.",
                    "Startup Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            catch
            {
                // WPF text rendering may be broken — write crash log to disk as last resort
                try
                {
                    var logPath = Path.Combine(Path.GetTempPath(), "LEQControlPanel_crash.log");
                    File.WriteAllText(logPath, $"[{DateTime.Now:O}] Critical error: {exception}");
                }
                catch { /* Nothing left to try */ }
            }
        };

        DispatcherUnhandledException += (sender, args) =>
        {
            (Application.Current.MainWindow as MainWindow)?.DisposeTrayIcon();
            try { if (_instanceMutex != null && _mutexOwned) _instanceMutex.ReleaseMutex(); } catch { }
            Debug.WriteLine($"[LEQControlPanel] UI thread error: {args.Exception}");
            try
            {
                System.Windows.MessageBox.Show(
                    "An unexpected error occurred. Please restart LEQ Control Panel.\n\n" +
                    "If this continues, submit a diagnostic report.",
                    "Application Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            catch
            {
                try
                {
                    var logPath = Path.Combine(Path.GetTempPath(), "LEQControlPanel_crash.log");
                    File.WriteAllText(logPath, $"[{DateTime.Now:O}] UI thread error: {args.Exception}");
                }
                catch { }
            }
            args.Handled = true; // Prevent app from crashing
        };

        TaskScheduler.UnobservedTaskException += (sender, args) =>
        {
            Debug.WriteLine($"[LEQControlPanel] Unobserved task exception: {args.Exception}");
            args.SetObserved();
        };

        base.OnStartup(e);

        // Clean up .old file from a previous successful update (fire-and-forget)
        _ = UpdateService.CleanupOldExecutable();

        try
        {
            // Check for command line arguments
            if (e.Args.Length > 0)
            {
                string arg = e.Args[0].ToLowerInvariant();

                if (arg == "-silent" || arg == "-toggle")
                {
                    // Run LEQ toggle logic headless
                    await RunHeadlessLEQToggle();
                    // Shutdown after headless operation
                    Shutdown();
                    return;
                }
            }

            // Single-instance check (GUI mode only — headless toggle is allowed to run concurrently)
            _instanceMutex = new Mutex(true, MutexName, out bool createdNew);
            if (!createdNew)
            {
                _instanceMutex.Dispose();
                _instanceMutex = null;
                // Another GUI instance is already running — bring it to the foreground
                ActivateExistingInstance();
                Shutdown();
                return;
            }
            _mutexOwned = true;

            // Show splash as the very first window — this gets on screen ASAP
            var splash = new SplashWindow();
            splash.Show();

            // Yield to let WPF render the splash before doing heavy work
            await Task.Delay(1);

            var mainWindow = new MainWindow();

            if (mainWindow._startMinimized)
            {
                splash.Close();
            }
            else
            {
                mainWindow.SplashScreen = splash;
                mainWindow.Opacity = 0;
            }

            mainWindow.Show();

            // Check for updates (fire and forget, don't block startup)
            _ = CheckForUpdatesAsync();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[LEQControlPanel] Application startup failed: {ex}");
            StyledMessageBox.SafeShowError(
                "LEQ Control Panel failed to start.\n\n" +
                "Please restart the application. If this continues, submit a diagnostic report.",
                "Startup Failed");
            Shutdown();
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        // Safety net: ensure tray icon is cleaned up on any exit path
        (MainWindow as MainWindow)?.DisposeTrayIcon();

        if (_instanceMutex != null && _mutexOwned)
        {
            try { _instanceMutex.ReleaseMutex(); } catch { }
            try { _instanceMutex.Dispose(); } catch { }
            _instanceMutex = null;
        }
        base.OnExit(e);
    }

    private static void ActivateExistingInstance()
    {
        using var currentProcess = Process.GetCurrentProcess();
        var processes = Process.GetProcessesByName(currentProcess.ProcessName);
        try
        {
            foreach (var process in processes)
            {
                if (process.Id != currentProcess.Id && process.MainWindowHandle != IntPtr.Zero)
                {
                    if (IsIconic(process.MainWindowHandle))
                        ShowWindow(process.MainWindowHandle, SW_RESTORE);
                    SetForegroundWindow(process.MainWindowHandle);
                    break;
                }
            }
        }
        finally
        {
            foreach (var process in processes)
                process.Dispose();
        }
    }

    private async Task CheckForUpdatesAsync()
    {
        // Wait a bit for app to fully initialize
        await Task.Delay(2000);

        var updateCheck = await UpdateChecker.CheckForUpdateAsync();

        if (updateCheck.HasValue && updateCheck.Value.UpdateAvailable)
        {
            var (_, newVersion, downloadUrl, sha256Hash) = updateCheck.Value;

            // Show update notification on UI thread
            var dialogResult = await Dispatcher.InvokeAsync(() =>
                StyledMessageBox.Show(
                    $"A new version is available!\n\n" +
                    $"Current: {Assembly.GetExecutingAssembly().GetName().Version?.ToString(3)}\n" +
                    $"Latest: {newVersion}\n\n" +
                    $"Would you like to download and install it now?",
                    "Update Available",
                    StyledMessageBoxButton.YesNo,
                    StyledMessageBoxIcon.Info
                ));

            if (dialogResult == MessageBoxResult.Yes)
            {
                bool relaunching = await UpdateService.DownloadAndSwapAsync(
                    downloadUrl, newVersion, MainWindow, sha256Hash);

                if (relaunching)
                {
                    if (MainWindow is MainWindow mw)
                        mw.ExitApplication();
                    else
                        Shutdown();
                }
            }
        }
    }

    private async Task RunHeadlessLEQToggle()
    {
        try
        {
            var audioService = new Services.AudioService();

            // Get devices
            var devices = await audioService.GetDevicesAsync();
            if (devices == null || devices.Count == 0)
            {
                return;
            }

            // Read saved device preference from registry
            string? savedGuid = null;
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\LEQControlPanel");
                savedGuid = key?.GetValue("LastSelectedDeviceGuid", "") as string;
            }
            catch { /* Registry read is best-effort — fallback to default device selection */ }

            // Find device: saved preference > first LEQ-configured > first available
            var preferredDevice =
                (!string.IsNullOrEmpty(savedGuid) ? devices.FirstOrDefault(d => d.Guid == savedGuid) : null)
                ?? devices.FirstOrDefault(d => d.LeqConfigured)
                ?? devices.First();
            if (preferredDevice != null)
            {
                // Toggle LEQ
                await audioService.ToggleLeqAsync(preferredDevice.Guid);
            }
        }
        catch (Exception)
        {
#if DEBUG
            Debug.WriteLine("Headless LEQ toggle failed");
#endif
            // Silent fail for headless mode - expected when no device available
        }
    }
}

