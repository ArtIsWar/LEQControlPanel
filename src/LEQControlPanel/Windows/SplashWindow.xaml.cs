// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System;
using System.Windows;
using System.Windows.Media.Animation;

namespace LEQControlPanel.Windows;

public partial class SplashWindow : Window
{
    public SplashWindow()
    {
        InitializeComponent();
        Loaded += SplashWindow_Loaded;
    }

    private void SplashWindow_Loaded(object sender, RoutedEventArgs e)
    {
        var glowPulse = (Storyboard)FindResource("GlowPulseAnimation");
        glowPulse.Begin();

        var barShimmer = (Storyboard)FindResource("BarShimmerAnimation");
        barShimmer.Begin();
    }

    public void FadeOutAndClose()
    {
        var fadeOut = (Storyboard)FindResource("FadeOutAnimation");
        fadeOut.Completed += (_, _) => Close();
        fadeOut.Begin();
    }
}
