// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System;

namespace LEQControlPanel;

public static class Program
{
    [STAThread]
    public static void Main()
    {
        var app = new App();
        app.InitializeComponent();
        app.Run();
    }
}
