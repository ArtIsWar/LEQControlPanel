// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using System.Windows.Navigation;

using LEQControlPanel.Dialogs;

namespace LEQControlPanel.Windows;

public partial class AboutWindow : Window
{
    public AboutWindow()
    {
        InitializeComponent();
        VersionText.Text = $"Version {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "1.0.0"}";
    }

    private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true })?.Dispose();
            e.Handled = true;
        }
        catch (Exception ex) { Debug.WriteLine($"[AboutWindow] Hyperlink navigation failed: {ex.Message}"); }
    }

    private void ViewLicense_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // Load LICENSE from embedded resource
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            using var stream = assembly.GetManifestResourceStream("LEQControlPanel.LICENSE");
            if (stream != null)
            {
                using var reader = new StreamReader(stream);
                string licenseText = reader.ReadToEnd();

                // Show license in a message box since we can't write to temp file in single-file mode
                StyledMessageBox.ShowInfo(licenseText, "License (GPL v3.0)");
            }
            else
            {
                StyledMessageBox.ShowInfo("License file not found in resources.\nReleased under GPL v3.0.", "License Info");
            }
        }
        catch (Exception ex)
        {
            StyledMessageBox.SafeShowError($"Could not load license: {ex.Message}", "License Error");
        }
    }

    private void VisitWebsite_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo("https://artiswar.io/arttunekit") { UseShellExecute = true })?.Dispose();
        }
        catch (Exception ex)
        {
            StyledMessageBox.SafeShowError($"Could not open website: {ex.Message}", "Error");
        }
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        DragMove();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
