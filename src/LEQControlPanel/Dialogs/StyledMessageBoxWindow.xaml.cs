// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;

namespace LEQControlPanel.Dialogs;

public partial class StyledMessageBoxWindow : Window
{
    public MessageBoxResult Result { get; private set; } = MessageBoxResult.None;

    public string? YesButtonText { set { if (value != null) YesButton.Content = value; } }
    public string? NoButtonText { set { if (value != null) NoButton.Content = value; } }

    public StyledMessageBoxWindow(Inline[] inlines, string title, StyledMessageBoxButton buttons, StyledMessageBoxIcon icon)
        : this((string?)null, title, buttons, icon)
    {
        MessageText.Text = null;
        MessageText.Inlines.Clear();
        MessageText.Inlines.AddRange(inlines);
    }

    public StyledMessageBoxWindow(string? message, string title, StyledMessageBoxButton buttons, StyledMessageBoxIcon icon)
    {
        InitializeComponent();

        MessageText.Text = message ?? "";
        TitleText.Text = string.IsNullOrEmpty(title) ? "LEQ Control Panel" : title;

        // Set icon
        switch (icon)
        {
            case StyledMessageBoxIcon.Info:
                IconText.Text = "\u2139\uFE0F";
                break;
            case StyledMessageBoxIcon.Warning:
                IconText.Text = "\u26A0\uFE0F";
                break;
            case StyledMessageBoxIcon.Error:
                IconText.Text = "\u274C";
                break;
            case StyledMessageBoxIcon.Question:
                IconText.Text = "\u2753";
                break;
            case StyledMessageBoxIcon.Danger:
                IconText.Text = "\u26A0\uFE0F";
                IconText.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0xD9, 0x44, 0x44));
                break;
        }

        // Configure buttons
        switch (buttons)
        {
            case StyledMessageBoxButton.OK:
                OkButton.Visibility = Visibility.Visible;
                break;
            case StyledMessageBoxButton.YesNo:
                YesButton.Visibility = Visibility.Visible;
                NoButton.Visibility = Visibility.Visible;
                NoButton.Style = (Style)FindResource("SecondaryButtonStyle");
                break;
            case StyledMessageBoxButton.YesNoCancel:
                YesButton.Visibility = Visibility.Visible;
                NoButton.Visibility = Visibility.Visible;
                CancelButton.Visibility = Visibility.Visible;
                NoButton.Style = (Style)FindResource("SecondaryButtonStyle");
                CancelButton.Style = (Style)FindResource("SecondaryButtonStyle");
                break;
        }

        // Apply danger styling to confirm button
        if (icon == StyledMessageBoxIcon.Danger)
        {
            YesButton.Style = (Style)FindResource("DangerButtonStyle");
            OkButton.Style = (Style)FindResource("DangerButtonStyle");
        }
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => DragMove();

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        Result = MessageBoxResult.OK;
        Close();
    }

    private void YesButton_Click(object sender, RoutedEventArgs e)
    {
        Result = MessageBoxResult.Yes;
        Close();
    }

    private void NoButton_Click(object sender, RoutedEventArgs e)
    {
        Result = MessageBoxResult.No;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        Result = MessageBoxResult.Cancel;
        Close();
    }
}
