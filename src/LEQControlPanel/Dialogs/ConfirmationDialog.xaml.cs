// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System.Windows;

namespace LEQControlPanel.Dialogs;

// public: required by WPF-generated partial class ConfirmationDialog
public enum CloseBehavior
{
    MinimizeToTray,
    ExitApplication
}

public partial class ConfirmationDialog : Window
{
    public CloseBehavior SelectedBehavior { get; private set; } = CloseBehavior.ExitApplication;
    public bool RememberChoice { get; private set; }

    public ConfirmationDialog()
    {
        InitializeComponent();
        this.Icon = null; // Prevent inheriting heavy 196KB icon from owner window
    }

    public ConfirmationDialog(CloseBehavior currentBehavior) : this()
    {
        // Pre-select the current behavior
        switch (currentBehavior)
        {
            case CloseBehavior.MinimizeToTray:
                OptionMinimizeToTray.IsChecked = true;
                break;
            case CloseBehavior.ExitApplication:
                OptionExitApplication.IsChecked = true;
                break;
        }
    }

    private void ConfirmButton_Click(object sender, RoutedEventArgs e)
    {
        if (OptionMinimizeToTray.IsChecked == true)
            SelectedBehavior = CloseBehavior.MinimizeToTray;
        else
            SelectedBehavior = CloseBehavior.ExitApplication;

        RememberChoice = RememberChoiceCheckbox.IsChecked == true;
        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
