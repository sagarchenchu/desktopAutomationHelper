using System;
using System.Collections.Generic;
using DesktopAutomationDriver.Models;
using DesktopAutomationDriver.Models.Request;
using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services;

public sealed class SearchRoot
{
    public AutomationElement? UiaElement { get; set; }
    public IntPtr? Hwnd { get; set; }

    public bool IsValid => (UiaElement != null) ^ (Hwnd != null);

    public void Validate()
    {
        if (!IsValid)
        {
            throw new InvalidOperationException("SearchRoot must have exactly one of UiaElement or Hwnd set.");
        }
    }
}

public sealed class ElementSnapshot
{
    public string Name { get; set; } = "";
    public string AutomationId { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string ControlType { get; set; } = "";
    public bool Enabled { get; set; }
    public bool Visible { get; set; }
    public System.Drawing.Rectangle Rectangle { get; set; }
}

public interface IElementBackend
{
    string Name { get; }

    IReadOnlyList<ElementCandidate> FindElements(
        SearchRoot root,
        UiLocator locator,
        ResolveOptions options);

    ElementSnapshot GetSnapshot(object nativeElement);

    bool Click(object nativeElement);
    bool SetFocus(object nativeElement);
    bool IsEnabled(object nativeElement);
    bool IsVisible(object nativeElement);
    string GetText(object nativeElement);
}

public sealed class UiaElementBackend : IElementBackend
{
    public string Name => "uia";

    public IReadOnlyList<ElementCandidate> FindElements(
        SearchRoot root,
        UiLocator locator,
        ResolveOptions options)
    {
        // TODO: Full implementation of UIA candidate lookup based on the resolved options.
        throw new NotImplementedException("UIA backend find is not fully implemented yet.");
    }

    public ElementSnapshot GetSnapshot(object nativeElement)
    {
        // TODO: Map native UIA element to element snapshot
        throw new NotImplementedException("UIA backend snapshot is not implemented.");
    }

    public bool Click(object nativeElement) => throw new NotImplementedException();
    public bool SetFocus(object nativeElement) => throw new NotImplementedException();
    public bool IsEnabled(object nativeElement) => throw new NotImplementedException();
    public bool IsVisible(object nativeElement) => throw new NotImplementedException();
    public string GetText(object nativeElement) => throw new NotImplementedException();
}

public sealed class Win32ElementBackend : IElementBackend
{
    public string Name => "win32";

    public IReadOnlyList<ElementCandidate> FindElements(
        SearchRoot root,
        UiLocator locator,
        ResolveOptions options)
    {
        // TODO: HWND enumeration, GetWindowText, GetClassName, GetDlgCtrlID, EnumChildWindows
        throw new NotImplementedException("Win32 backend find is not implemented.");
    }

    public ElementSnapshot GetSnapshot(object nativeElement)
    {
        throw new NotImplementedException();
    }

    public bool Click(object nativeElement) => throw new NotImplementedException();
    public bool SetFocus(object nativeElement) => throw new NotImplementedException();
    public bool IsEnabled(object nativeElement) => throw new NotImplementedException();
    public bool IsVisible(object nativeElement) => throw new NotImplementedException();
    public string GetText(object nativeElement) => throw new NotImplementedException();
}

public sealed class HybridElementBackend : IElementBackend
{
    public string Name => "hybrid";

    public IReadOnlyList<ElementCandidate> FindElements(
        SearchRoot root,
        UiLocator locator,
        ResolveOptions options)
    {
        // TODO: Try UIA first, fallback to Win32
        throw new NotImplementedException("Hybrid backend find is not implemented.");
    }

    public ElementSnapshot GetSnapshot(object nativeElement)
    {
        throw new NotImplementedException();
    }

    public bool Click(object nativeElement) => throw new NotImplementedException();
    public bool SetFocus(object nativeElement) => throw new NotImplementedException();
    public bool IsEnabled(object nativeElement) => throw new NotImplementedException();
    public bool IsVisible(object nativeElement) => throw new NotImplementedException();
    public string GetText(object nativeElement) => throw new NotImplementedException();
}
