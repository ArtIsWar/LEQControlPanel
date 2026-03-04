// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace LEQControlPanel.Utilities;

/// <summary>
/// Custom renderer for dark-themed system tray context menu
/// </summary>
internal class DarkTrayMenuRenderer : ToolStripProfessionalRenderer
{
    // Color scheme matching main app theme
    private static readonly Color BackgroundColor = Color.FromArgb(26, 26, 26);      // #1A1A1A
    private static readonly Color HoverColor = Color.FromArgb(255, 215, 0);          // #FFD700
    private static readonly Color TextColor = Color.White;
    private static readonly Color DisabledTextColor = Color.FromArgb(128, 128, 128); // Gray
    private static readonly Color SeparatorColor = Color.FromArgb(60, 60, 60);       // Subtle gray
    private static readonly Color BorderColor = Color.FromArgb(60, 60, 60);
    private static readonly Color CheckmarkColor = Color.FromArgb(255, 215, 0);      // Yellow

    public DarkTrayMenuRenderer() : base(new DarkColorTable()) { }

    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
    {
        if (!e.Item.Enabled)
        {
            // Disabled items - solid background
            e.Graphics.FillRectangle(new SolidBrush(BackgroundColor), e.Item.ContentRectangle);
            return;
        }

        if (e.Item.Selected)
        {
            // Hover state - yellow background with slight transparency
            using (var brush = new SolidBrush(Color.FromArgb(200, HoverColor)))
            {
                e.Graphics.FillRectangle(brush, e.Item.ContentRectangle);
            }
        }
        else
        {
            // Normal state
            e.Graphics.FillRectangle(new SolidBrush(BackgroundColor), e.Item.ContentRectangle);
        }
    }

    protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
    {
        if (e.Item is ToolStripMenuItem menuItem)
        {
            // Set text color based on state
            if (!e.Item.Enabled)
            {
                e.TextColor = DisabledTextColor;
            }
            else if (e.Item.Selected)
            {
                e.TextColor = Color.Black; // Black text on yellow hover
            }
            else
            {
                e.TextColor = TextColor;
            }
        }

        // Let base renderer handle the actual drawing
        base.OnRenderItemText(e);
    }

    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
    {
        // Draw horizontal line separator
        int y = e.Item.Height / 2;
        using (var pen = new Pen(SeparatorColor))
        {
            e.Graphics.DrawLine(pen,
                e.Item.ContentRectangle.Left + 25, y,
                e.Item.ContentRectangle.Right - 5, y);
        }
    }

    protected override void OnRenderItemCheck(ToolStripItemImageRenderEventArgs e)
    {
        // Draw checkmark for checked menu items - larger and better positioned
        var rect = new Rectangle(
            e.Item.ContentRectangle.Left + 3,    // Slightly right of left edge
            e.Item.ContentRectangle.Top + 4,     // Better vertical centering
            20, 20);                              // Even larger size

        // Draw checkmark background
        e.Graphics.FillRectangle(new SolidBrush(BackgroundColor), rect);

        // Draw checkmark symbol with anti-aliasing
        using (var pen = new Pen(CheckmarkColor, 2.5f))  // Thicker line
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            e.Graphics.DrawLines(pen, new Point[]
            {
                new Point(rect.Left + 4, rect.Top + 10),
                new Point(rect.Left + 9, rect.Top + 15),
                new Point(rect.Left + 16, rect.Top + 6)
            });
        }
    }

    protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
    {
        // Draw submenu arrow (if needed in future)
        e.ArrowColor = e.Item?.Enabled == true ? TextColor : DisabledTextColor;
        base.OnRenderArrow(e);
    }
}

/// <summary>
/// Custom color table for menu rendering
/// </summary>
internal class DarkColorTable : ProfessionalColorTable
{
    public override Color MenuItemSelected => Color.FromArgb(255, 215, 0);
    public override Color MenuItemBorder => Color.FromArgb(60, 60, 60);
    public override Color MenuBorder => Color.FromArgb(60, 60, 60);
    public override Color ImageMarginGradientBegin => Color.FromArgb(26, 26, 26);
    public override Color ImageMarginGradientMiddle => Color.FromArgb(26, 26, 26);
    public override Color ImageMarginGradientEnd => Color.FromArgb(26, 26, 26);
    public override Color MenuItemSelectedGradientBegin => Color.FromArgb(255, 215, 0);
    public override Color MenuItemSelectedGradientEnd => Color.FromArgb(255, 215, 0);
    public override Color MenuStripGradientBegin => Color.FromArgb(26, 26, 26);
    public override Color MenuStripGradientEnd => Color.FromArgb(26, 26, 26);
    public override Color ToolStripDropDownBackground => Color.FromArgb(26, 26, 26);
}
