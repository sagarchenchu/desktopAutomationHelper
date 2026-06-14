using System;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution.ControlActionAdapters;

public sealed class TreeItemActionAdapter : IControlActionAdapter
{
    public bool CanHandle(ElementSnapshot snapshot, string operation)
    {
        return snapshot.ControlType == "TreeItem" && (operation == "select" || operation == "click" || operation == "expandtreeitem" || operation == "collapsetreeitem");
    }

    public object Execute(AutomationElement element, UiRequest request)
    {
        var treeItem = element.AsTreeItem();
        if (request.Operation == "expandtreeitem")
        {
            treeItem.Expand();
            return new { success = true, state = "expanded" };
        }
        if (request.Operation == "collapsetreeitem")
        {
            treeItem.Collapse();
            return new { success = true, state = "collapsed" };
        }
        treeItem.Select();
        return new { success = true, strategy = "treeitem-select" };
    }
}
