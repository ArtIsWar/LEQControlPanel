// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System.Windows;

namespace LEQControlPanel.Dialogs;

public partial class EapoWarningDialog : Window
{
    public bool DontAskAgain => DontAskAgainCheck.IsChecked == true;
    public bool Proceed { get; private set; }

    public EapoWarningDialog()
    {
        InitializeComponent();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        Proceed = false;
        DialogResult = false;
        Close();
    }

    private void ProceedButton_Click(object sender, RoutedEventArgs e)
    {
        Proceed = true;
        DialogResult = true;
        Close();
    }
}
