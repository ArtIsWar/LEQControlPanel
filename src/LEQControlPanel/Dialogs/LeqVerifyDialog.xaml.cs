// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System.Windows;
using System.Windows.Input;

namespace LEQControlPanel.Dialogs;

public partial class LeqVerifyDialog : Window
{
    public MessageBoxResult Result { get; private set; } = MessageBoxResult.None;

    public LeqVerifyDialog()
    {
        InitializeComponent();
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => DragMove();

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
}
