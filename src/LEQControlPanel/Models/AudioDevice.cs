// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

namespace LEQControlPanel.Models;

/// <summary>
/// Represents an audio device with LEQ configuration information.
/// </summary>
internal class AudioDevice
{
    /// <summary>
    /// The friendly name of the audio device.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// The interface name (e.g., "Speakers (Realtek Audio)").
    /// </summary>
    public string? InterfaceName { get; set; }

    /// <summary>
    /// The device GUID.
    /// </summary>
    public string Guid { get; set; } = string.Empty;

    /// <summary>
    /// Whether the device supports audio enhancements.
    /// </summary>
    public bool SupportsEnhancement { get; set; }

    /// <summary>
    /// Whether the device has LFX/GFX APO slots in FxProperties.
    /// </summary>
    public bool HasLfxGfx { get; set; }

    /// <summary>
    /// Whether the device has CompositeFX keys blocking LFX/GFX slots.
    /// </summary>
    public bool HasCompositeFx { get; set; }

    /// <summary>
    /// Whether LEQ is configured on this device.
    /// </summary>
    public bool LeqConfigured { get; set; }

    /// <summary>
    /// Whether Loudness Equalization is currently enabled.
    /// </summary>
    public bool LoudnessEnabled { get; set; }

    /// <summary>
    /// The LEQ release time value (2-7).
    /// </summary>
    public int ReleaseTime { get; set; }

    /// <summary>
    /// The E-APO status string.
    /// </summary>
    public string EapoStatus { get; set; } = string.Empty;

    /// <summary>
    /// Whether the E-APO child configuration is broken.
    /// </summary>
    public bool EapoChildBroken { get; set; }

    /// <summary>
    /// Number of audio channels.
    /// </summary>
    public int Channels { get; set; }

    /// <summary>
    /// Audio bit depth.
    /// </summary>
    public int BitDepth { get; set; }

    /// <summary>
    /// Audio sample rate in Hz.
    /// </summary>
    public int SampleRate { get; set; }

    /// <summary>
    /// Whether the device has a release time registry key.
    /// </summary>
    public bool HasReleaseTimeKey { get; set; }
}
