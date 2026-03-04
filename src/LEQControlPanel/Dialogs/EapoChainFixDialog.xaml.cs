// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System.Windows;

namespace LEQControlPanel.Dialogs;

public partial class EapoChainFixDialog : Window
{
    public bool Proceed { get; private set; }

    public EapoChainFixDialog(string friendlyName, string interfaceName)
    {
        InitializeComponent();

        FriendlyNameText.Text = friendlyName ?? "Unknown Device";
        InterfaceNameText.Text = interfaceName ?? "";

        if (string.IsNullOrEmpty(interfaceName))
        {
            InterfaceNameText.Visibility = Visibility.Collapsed;
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        Proceed = false;
        DialogResult = false;
        Close();
    }

    private void CleanInstallButton_Click(object sender, RoutedEventArgs e)
    {
        Proceed = true;
        DialogResult = true;
        Close();
    }
}
