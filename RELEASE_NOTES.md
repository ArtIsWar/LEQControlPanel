# LEQ Control Panel v1.0.0

First public release. A standalone Windows tool for registry-level Loudness Equalization control and installation.

## Highlights

- **Instant LEQ toggle** with no audio service restart required
- **Release time control** (2-7 range) with real-time Windows UI sync
- **LEQ installation** on devices that lack native enhancement support, with clean install and uninstall options
- **Multi-device support** with automatic state detection and device change monitoring
- **Registry health checks** detect broken COM CLSID registrations and per-device CompositeFX keys blocking LEQ; the Install button turns red with a one-click fix
- **Post-install verification** opens a guided dialog alongside Windows Sound Panel to confirm LEQ is active, with Spatial Audio compatibility warnings
- **E-APO awareness** detects Equalizer APO and adjusts behavior — gates LEQ installation on E-APO device configuration, offers chain repair when conflicts are detected, and shows a "Get E-APO" button when E-APO is not installed
- **Smart device reset** adapts to device type: Voicemeeter devices excluded, VB-Audio and Hi-Fi virtual cables warn of full driver removal, physical devices list affected sibling endpoints before proceeding
- **Audio format display** shows speaker layout (Stereo, 5.1, 7.1, etc.) and format details (bit depth, sample rate) in the device info strip
- **Activity log** collapsible console panel logs all actions in real time with timestamps, color-coded warnings/errors
- **System tray integration** with quick toggle, device info, and settings access
- **Headless CLI mode** (`-silent` / `-toggle`) for global hotkey binding and scripting
- **ArtTuneKit integration** detects ArtTuneKit and defers device management when installed; shows an overlay on Art Tune devices with a button to launch ArtTuneKit directly
- **Built-in auto-update** checks GitHub releases on startup with manual check from Settings menu, SHA256 integrity verification for downloaded updates
- **Settings:** run at startup, start minimized, always-on-top, desktop shortcut, configurable close behavior (minimize to tray / exit / ask)
- **Utilities:** Open Windows Sound settings and restart Windows Audio service from within the app

## CLI Usage

| Flag | Behavior |
|------|----------|
| *(none)* | Launch the full GUI |
| `-silent` | Toggle LEQ on the preferred device and exit immediately (no window) |
| `-toggle` | Same as `-silent` |

In headless mode, the application uses the device last selected in the GUI (saved in the registry), falling back to the first device with LEQ configured or the first available device. It toggles LEQ and exits. No GUI is displayed. The process can run concurrently with an open GUI instance.

To bind a global hotkey, create a Windows shortcut targeting `LEQControlPanel.exe -toggle` and assign a key combination in the shortcut's properties.

## System Requirements

- Windows 10 or 11 (x64)
- Administrator privileges (required for registry access)
- No .NET runtime installation needed (self-contained single-file executable)

## Building from Source

Requires [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) (Windows x64):

```bash
dotnet publish src/LEQControlPanel/LEQControlPanel.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=false
```

## Credits

- Original concept: [Falcosc/enable-loudness-equalisation](https://github.com/Falcosc/enable-loudness-equalisation)
- Extended implementation: [ArtIsWar](https://github.com/ArtIsWar)

## License

GNU General Public License v3.0
