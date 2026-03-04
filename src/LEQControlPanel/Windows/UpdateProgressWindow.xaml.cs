// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System;
using System.Threading;
using System.Windows;
using System.Windows.Input;

namespace LEQControlPanel.Windows;

public partial class UpdateProgressWindow : Window
{
    public CancellationTokenSource Cts { get; } = new();

    public UpdateProgressWindow()
    {
        InitializeComponent();
    }

    public void SetVersion(string version)
    {
        VersionText.Text = $"Downloading v{version}...";
    }

    public void UpdateProgress(double percent, string statusText)
    {
        // Update the fill width proportionally to the parent border
        ProgressFill.Width = Math.Max(0, (ActualWidth - 40) * (percent / 100.0)); // 40 = left+right margin
        StatusText.Text = statusText;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        Cts.Cancel();
        Close();
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        DragMove();
    }

    protected override void OnClosed(EventArgs e)
    {
        if (!Cts.IsCancellationRequested)
            Cts.Cancel();
        Cts.Dispose();
        base.OnClosed(e);
    }
}
