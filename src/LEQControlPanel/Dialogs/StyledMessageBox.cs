// LEQ Control Panel — Copyright (c) 2025-2026 ArtIsWar LLC
// Licensed under GPL-3.0. See LICENSE file for details.

using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;

namespace LEQControlPanel.Dialogs;

// public: required by WPF-generated partial class StyledMessageBoxWindow
public enum StyledMessageBoxButton { OK, YesNo, YesNoCancel }
public enum StyledMessageBoxIcon { Info, Warning, Error, Question, Danger }

internal static class StyledMessageBox
{
    public static MessageBoxResult Show(string message, string title = "",
        StyledMessageBoxButton buttons = StyledMessageBoxButton.OK,
        StyledMessageBoxIcon icon = StyledMessageBoxIcon.Info)
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.CheckAccess())
            return dispatcher.Invoke(() => Show(message, title, buttons, icon));

        var dialog = new StyledMessageBoxWindow(message, title, buttons, icon);
        ApplyOwner(dialog);
        dialog.ShowDialog();
        return dialog.Result;
    }

    public static MessageBoxResult Show(System.Windows.Documents.Inline[] inlines, string title = "",
        StyledMessageBoxButton buttons = StyledMessageBoxButton.OK,
        StyledMessageBoxIcon icon = StyledMessageBoxIcon.Info)
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.CheckAccess())
            return dispatcher.Invoke(() => Show(inlines, title, buttons, icon));

        var dialog = new StyledMessageBoxWindow(inlines, title, buttons, icon);
        ApplyOwner(dialog);
        dialog.ShowDialog();
        return dialog.Result;
    }

    // Convenience methods
    public static void ShowInfo(string message, string title = "Information")
        => Show(message, title, StyledMessageBoxButton.OK, StyledMessageBoxIcon.Info);

    public static void ShowWarning(string message, string title = "Warning")
        => Show(message, title, StyledMessageBoxButton.OK, StyledMessageBoxIcon.Warning);

    public static void ShowError(string message, string title = "Error")
        => Show(message, title, StyledMessageBoxButton.OK, StyledMessageBoxIcon.Error);

    public static MessageBoxResult ShowYesNo(string message, string title = "Confirm")
        => Show(message, title, StyledMessageBoxButton.YesNo, StyledMessageBoxIcon.Question);

    public static MessageBoxResult ShowYesNo(System.Windows.Documents.Inline[] inlines, string title = "Confirm")
        => Show(inlines, title, StyledMessageBoxButton.YesNo, StyledMessageBoxIcon.Question);

    public static MessageBoxResult ShowYesNo(string message, string title = "Confirm",
        string yesText = "Yes", string noText = "No")
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.CheckAccess())
            return dispatcher.Invoke(() => ShowYesNo(message, title, yesText, noText));

        var dialog = new StyledMessageBoxWindow(message, title, StyledMessageBoxButton.YesNo, StyledMessageBoxIcon.Question);
        dialog.YesButtonText = yesText;
        dialog.NoButtonText = noText;
        ApplyOwner(dialog);
        dialog.ShowDialog();
        return dialog.Result;
    }

    public static MessageBoxResult ShowYesNoCancel(string message, string title = "Confirm")
        => Show(message, title, StyledMessageBoxButton.YesNoCancel, StyledMessageBoxIcon.Question);

    public static MessageBoxResult ShowDanger(string message, string title = "Confirm",
        string confirmText = "Remove", string cancelText = "Cancel")
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.CheckAccess())
            return dispatcher.Invoke(() => ShowDanger(message, title, confirmText, cancelText));

        var dialog = new StyledMessageBoxWindow(message, title, StyledMessageBoxButton.YesNo, StyledMessageBoxIcon.Danger);
        dialog.YesButtonText = confirmText;
        dialog.NoButtonText = cancelText;
        ApplyOwner(dialog);
        dialog.ShowDialog();
        return dialog.Result;
    }

    // Safe wrappers — catch dialog failures (e.g. WPF text rendering crash after reboot)
    // and fall back to Debug.WriteLine so exception handlers don't cascade.

    public static MessageBoxResult SafeShow(string message, string title = "",
        StyledMessageBoxButton buttons = StyledMessageBoxButton.OK,
        StyledMessageBoxIcon icon = StyledMessageBoxIcon.Info)
    {
        try { return Show(message, title, buttons, icon); }
        catch (Exception ex)
        {
            Debug.WriteLine($"[StyledMessageBox] Dialog failed: {ex.Message} — Original: {title}: {message}");
            return MessageBoxResult.None;
        }
    }

    public static void SafeShowInfo(string message, string title = "Information")
    {
        try { ShowInfo(message, title); }
        catch (Exception ex) { Debug.WriteLine($"[StyledMessageBox] Dialog failed: {ex.Message} — Original: {title}: {message}"); }
    }

    public static void SafeShowWarning(string message, string title = "Warning")
    {
        try { ShowWarning(message, title); }
        catch (Exception ex) { Debug.WriteLine($"[StyledMessageBox] Dialog failed: {ex.Message} — Original: {title}: {message}"); }
    }

    public static void SafeShowError(string message, string title = "Error")
    {
        try { ShowError(message, title); }
        catch (Exception ex) { Debug.WriteLine($"[StyledMessageBox] Dialog failed: {ex.Message} — Original: {title}: {message}"); }
    }

    private static Window? ResolveOwner()
    {
        Window? best = null;
        foreach (Window w in System.Windows.Application.Current.Windows)
        {
            if (w is StyledMessageBoxWindow) continue;
            if (w.Visibility != Visibility.Visible) continue;
            if (w.WindowState == WindowState.Minimized) continue;
            if (!w.IsLoaded) continue;

            if (w.IsActive)
                return w;

            best = w;
        }

        return best;
    }

    private static void ApplyOwner(Window dialog)
    {
        var owner = ResolveOwner();
        if (owner != null)
        {
            try
            {
                dialog.Owner = owner;
            }
            catch
            {
                dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
        }
        else
        {
            dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }
    }
}
