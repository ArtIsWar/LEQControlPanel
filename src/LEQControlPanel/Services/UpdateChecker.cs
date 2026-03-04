// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Reflection;

namespace LEQControlPanel.Services;

internal class UpdateChecker
{
    private const string GitHubOwner = "ArtIsWar";
    private const string GitHubRepo = "LEQControlPanel";
    private const string GitHubReleasesUrl = $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest";
    private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };

    public static async Task<(bool UpdateAvailable, string NewVersion, string DownloadUrl, string? Sha256Hash)?> CheckForUpdateAsync()
    {
        try
        {
            if (!_httpClient.DefaultRequestHeaders.UserAgent.Any())
            {
                _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("LEQControlPanel");
            }

            // Get current version from assembly
            var currentVersion = Assembly.GetExecutingAssembly()
                .GetName()
                .Version?
                .ToString(3) ?? "1.0.0"; // Format: Major.Minor.Build

#if DEBUG
            System.Diagnostics.Debug.WriteLine($"Current version: {currentVersion}");
#endif

            // Fetch latest release from GitHub Releases API
            var response = await _httpClient.GetStringAsync(GitHubReleasesUrl);
            using var doc = JsonDocument.Parse(response);
            var root = doc.RootElement;

            if (!root.TryGetProperty("tag_name", out var tagNameElement))
            {
                return null;
            }

            var tagName = tagNameElement.GetString() ?? string.Empty;
            if (tagName.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                tagName = tagName.Substring(1);
            }

            if (string.IsNullOrWhiteSpace(tagName))
            {
                return null;
            }

            string? downloadUrl = null;
            if (root.TryGetProperty("assets", out var assetsElement) &&
                assetsElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var asset in assetsElement.EnumerateArray())
                {
                    if (!asset.TryGetProperty("name", out var nameElement))
                        continue;

                    var assetName = nameElement.GetString() ?? string.Empty;
                    if (!assetName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (asset.TryGetProperty("browser_download_url", out var urlElement))
                    {
                        downloadUrl = urlElement.GetString();
                        if (!string.IsNullOrWhiteSpace(downloadUrl))
                            break;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(downloadUrl))
            {
                return null;
            }

            // Parse SHA256 hash from release body (format: "SHA256: <64-char-hex>")
            string? sha256Hash = null;
            if (root.TryGetProperty("body", out var bodyElement))
            {
                var body = bodyElement.GetString() ?? string.Empty;
                var match = Regex.Match(body, @"SHA256:\s*([0-9a-fA-F]{64})", RegexOptions.IgnoreCase);
                if (match.Success)
                    sha256Hash = match.Groups[1].Value;
            }

#if DEBUG
            System.Diagnostics.Debug.WriteLine($"Remote version: {tagName}, SHA256: {sha256Hash ?? "(not published)"}");
#endif

            // Compare versions
            var current = new Version(currentVersion);
            var remote = new Version(tagName);

            if (remote > current)
            {
                return (true, tagName, downloadUrl, sha256Hash);
            }

            return (false, tagName, downloadUrl, sha256Hash);
        }
        catch (Exception)
        {
#if DEBUG
            System.Diagnostics.Debug.WriteLine("Update check failed");
#endif
            // Silent fail - don't bother user if check fails
            return null;
        }
    }
}
