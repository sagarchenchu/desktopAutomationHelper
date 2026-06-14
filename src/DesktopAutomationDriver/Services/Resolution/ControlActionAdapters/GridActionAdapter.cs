using System;
using System.Collections.Generic;
using System.Linq;
using FlaUI.Core.AutomationElements;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution.ControlActionAdapters;

public sealed class GridActionAdapter : IControlActionAdapter
{
    public bool CanHandle(ElementSnapshot snapshot, string operation)
    {
        return (snapshot.ControlType == "DataGrid" || snapshot.ControlType == "Table") && 
               (operation == "clickgridcell" || operation == "doubleclickgridcell" || operation == "gettable" || operation == "gettabledata");
    }

    public object Execute(AutomationElement element, UiRequest request)
    {
        var grid = element.AsDataGridView();
        if (request.Operation == "gettable" || request.Operation == "gettabledata")
        {
            var rowsList = new List<List<string>>();
            foreach (var row in grid.Rows)
            {
                var cells = row.Cells.Select(c => c.Value ?? "").ToList();
                rowsList.Add(cells);
            }
            int colCount = grid.Rows.FirstOrDefault()?.Cells?.Length ?? 0;
            return new { success = true, rows = rowsList, rowCount = grid.Rows.Length, columnCount = colCount };
        }

        if (request.Index.HasValue && request.ColumnIndex.HasValue)
        {
            var rowIdx = request.Index.Value;
            var colIdx = request.ColumnIndex.Value;
            if (rowIdx >= 0 && rowIdx < grid.Rows.Length)
            {
                var row = grid.Rows[rowIdx];
                if (colIdx >= 0 && colIdx < row.Cells.Length)
                {
                    var cell = row.Cells[colIdx];
                    if (request.Operation == "doubleclickgridcell")
                    {
                        cell.DoubleClick();
                        return new { success = true, strategy = "physical-doubleclick-gridcell" };
                    }
                    cell.Click();
                    return new { success = true, strategy = "physical-click-gridcell" };
                }
            }
        }
        return new { success = false, message = "Row or column index out of bounds or not provided" };
    }
}
