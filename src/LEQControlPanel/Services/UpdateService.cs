// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using LEQControlPanel.Dialogs;
using LEQControlPanel.Windows;

namespace LEQControlPanel.Services;

internal static class UpdateService
{
    private static readonly HttpClient _downloadClient;

    static UpdateService()
    {
        _downloadClient = new HttpClient
        {
            Timeout = TimeSpan.FromMinutes(10)
        };
        _downloadClient.DefaultRequestHeaders.UserAgent.ParseAdd("LEQControlPanel");
    }

    /// <summary>
    /// Deletes LEQControlPanel.exe.old left over from a previous update.
    /// Retries with delay since the old process may still be exiting.
    /// Fire-and-forget — safe to call on every startup.
    /// </summary>
    public static async Task CleanupOldExecutable()
    {
        try
        {
            var exePath = Environment.ProcessPath;
            if (string.IsNullOrEmpty(exePath)) return;

            var oldPath = exePath + ".old";
            if (!File.Exists(oldPath)) return;

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    File.Delete(oldPath);
                    return; // Success
                }
                catch (IOException) when (attempt < 2)
                {
                    // File likely still locked by the exiting old process
                    await Task.Delay(1000);
                }
            }
        }
        catch (Exception ex)
        {
            // Best-effort; .old is inert, will be cleaned up next launch
            Debug.WriteLine($"[UpdateService] CleanupOldExecutable failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Downloads the new exe, swaps it in place of the running one, and relaunches.
    /// </summary>
    /// <returns>true if relaunch was initiated (caller should shut down);
    /// false if cancelled or failed (caller continues running).</returns>
    public static async Task<bool> DownloadAndSwapAsync(
        string downloadUrl,
        string newVersion,
        Window ownerWindow,
        string? expectedHash = null)
    {
        var exePath = Environment.ProcessPath;
        if (string.IsNullOrEmpty(exePath))
        {
            StyledMessageBox.ShowError(
                "Could not determine the application path.\n\nPlease update manually.",
                "Update Failed");
            return false;
        }

        var exeDir = Path.GetDirectoryName(exePath)!;
        var exeName = Path.GetFileName(exePath);
        var tempPath = Path.Combine(exeDir, exeName + ".new");
        var oldPath = exePath + ".old";

        // Show progress window
        var progressWindow = new UpdateProgressWindow();
        try
        {
            if (ownerWindow.IsLoaded)
                progressWindow.Owner = ownerWindow;
        }
        catch (Exception ex) { Debug.WriteLine($"[UpdateService] Could not set progress window owner: {ex.Message}"); }
        progressWindow.SetVersion(newVersion);
        progressWindow.Show();

        try
        {
            // --- Phase 1: Download to temp file ---
            using var response = await _downloadClient.GetAsync(
                downloadUrl,
                HttpCompletionOption.ResponseHeadersRead,
                progressWindow.Cts.Token);

            response.EnsureSuccessStatusCode();

            var totalBytes = response.Content.Headers.ContentLength;
            long downloadedBytes = 0;

            {
                await using var contentStream = await response.Content.ReadAsStreamAsync(
                    progressWindow.Cts.Token);
                await using var fileStream = new FileStream(
                    tempPath, FileMode.Create, FileAccess.Write, FileShare.None, 81920);

                var buffer = new byte[81920];
                int bytesRead;

                while ((bytesRead = await contentStream.ReadAsync(
                    buffer.AsMemory(0, buffer.Length),
                    progressWindow.Cts.Token)) > 0)
                {
                    await fileStream.WriteAsync(
                        buffer.AsMemory(0, bytesRead),
                        progressWindow.Cts.Token);

                    downloadedBytes += bytesRead;

                    double percent = totalBytes is > 0
                        ? (double)downloadedBytes / totalBytes.Value * 100.0
                        : 0;

                    string sizeText = totalBytes.HasValue
                        ? $"{FormatBytes(downloadedBytes)} / {FormatBytes(totalBytes.Value)}"
                        : FormatBytes(downloadedBytes);

                    progressWindow.UpdateProgress(percent, $"{percent:F0}% \u2014 {sizeText}");
                }
            }

            progressWindow.Close();

            // --- Phase 1b: SHA256 verification ---
            if (!string.IsNullOrEmpty(expectedHash))
            {
                using var sha256 = SHA256.Create();
                await using var hashStream = File.OpenRead(tempPath);
                var hashBytes = await sha256.ComputeHashAsync(hashStream);
                var actualHash = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();

                if (!string.Equals(actualHash, expectedHash, StringComparison.OrdinalIgnoreCase))
                {
                    TryCleanup(tempPath);
                    StyledMessageBox.ShowError(
                        "The downloaded update could not be verified and has been discarded.\n\n" +
                        "Please try again, or download manually from GitHub.",
                        "Update Verification Failed");
                    return false;
                }
            }

            // --- Phase 2: Swap ---
            if (File.Exists(oldPath))
                File.Delete(oldPath);

            // Rename running exe -> .old (Windows allows rename of running exe)
            File.Move(exePath, oldPath);

            // Move downloaded exe -> original name
            File.Move(tempPath, exePath);
        }
        catch (OperationCanceledException)
        {
            TryCleanup(tempPath);
            if (progressWindow.IsLoaded)
                progressWindow.Close();
            return false;
        }
        catch (Exception ex)
        {
            if (progressWindow.IsLoaded)
                progressWindow.Close();

            // Rollback: if we renamed the exe but failed before the new one was in place
            if (!File.Exists(exePath) && File.Exists(oldPath))
            {
                try { File.Move(oldPath, exePath); }
                catch (Exception rollbackEx) { Debug.WriteLine($"[UpdateService] Rollback failed: {rollbackEx.Message}"); }
            }
            TryCleanup(tempPath);

            StyledMessageBox.SafeShowError(
                $"Update failed:\n\n{ex.Message}\n\nThe application has not been modified.",
                "Update Failed");
            return false;
        }

        // --- Phase 3: Relaunch (swap already succeeded at this point) ---
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = exePath,
                UseShellExecute = true
            })?.Dispose();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[UpdateService] Relaunch failed: {ex}");
            StyledMessageBox.SafeShowError(
                $"Update was applied successfully, but the app couldn't restart:\n\n{ex.Message}\n\n" +
                "Please launch the application manually.",
                "Restart Failed");
        }

        return true; // Caller should shut down regardless — swap succeeded
    }

    private static void TryCleanup(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (Exception ex) { Debug.WriteLine($"[UpdateService] TryCleanup failed for {path}: {ex.Message}"); }
    }

    private static string FormatBytes(long bytes) => bytes switch
    {
        >= 1_073_741_824 => $"{bytes / 1_073_741_824.0:F1} GB",
        >= 1_048_576 => $"{bytes / 1_048_576.0:F1} MB",
        >= 1024 => $"{bytes / 1024.0:F1} KB",
        _ => $"{bytes} B"
    };
}
