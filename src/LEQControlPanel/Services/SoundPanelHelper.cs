// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System.Diagnostics;
using Microsoft.Win32;

namespace LEQControlPanel.Services;

/// <summary>
/// Helper for launching the Windows Sound control panel (mmsys.cpl).
/// </summary>
internal static class SoundPanelHelper
{
    private const string DeviceCplRegistryPath = @"Software\Microsoft\Multimedia\Audio\DeviceCpl";

    /// <summary>
    /// Opens the Windows Sound control panel (mmsys.cpl), hiding disconnected and disabled devices.
    /// </summary>
    public static void OpenSoundPanel()
    {
        HideDisconnectedAndDisabledDevices();
        Process.Start(new ProcessStartInfo("mmsys.cpl") { UseShellExecute = true })?.Dispose();
    }

    /// <summary>
    /// Sets registry values to hide disconnected and disabled devices in the Sound control panel.
    /// These are per-user settings stored in HKCU.
    /// </summary>
    private static void HideDisconnectedAndDisabledDevices()
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(DeviceCplRegistryPath);
            key?.SetValue("ShowDisconnectedDevices", 0, RegistryValueKind.DWord);
            key?.SetValue("ShowDisabledDevices", 0, RegistryValueKind.DWord);
        }
        catch
        {
            // Ignore if we can't write to registry - mmsys.cpl will still open
        }
    }
}
