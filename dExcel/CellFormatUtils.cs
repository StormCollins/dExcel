namespace dExcel;

using ExcelDna.Integration;
using static Microsoft.Office.Interop.Excel.XlRgbColor;
using Excel = Microsoft.Office.Interop.Excel;

public enum TableOrientation
{
    Both, Horizontal, Vertical,
}

public enum HorizontalCellAlignment
{
    Center, CenterAcrossSelection, Left, Right,
}

public enum VerticalCellAlignment
{
    Bottom, Center, CenterAcrossSelection, Top,
}

public static class CellFormatUtils
{
    private static readonly Color DeloitteGreen = Color.FromArgb(134, 188, 37);
    private static readonly Excel.XlBordersIndex[] BoundaryBorderIndices =
        {
                Excel.XlBordersIndex.xlEdgeBottom,
                Excel.XlBordersIndex.xlEdgeLeft,
                Excel.XlBordersIndex.xlEdgeRight,
                Excel.XlBordersIndex.xlEdgeTop,
            };

    private static readonly Excel.XlBordersIndex[] InteriorBorderIndices =
        {
                Excel.XlBordersIndex.xlInsideHorizontal,
                Excel.XlBordersIndex.xlInsideVertical,
            };

    private static readonly Dictionary<HorizontalCellAlignment, Excel.XlHAlign> HorizontalAlignment
        = new()
        {
            { HorizontalCellAlignment.Center, Excel.XlHAlign.xlHAlignCenter },
            { HorizontalCellAlignment.CenterAcrossSelection, Excel.XlHAlign.xlHAlignCenterAcrossSelection },
            { HorizontalCellAlignment.Left, Excel.XlHAlign.xlHAlignLeft },
            { HorizontalCellAlignment.Right, Excel.XlHAlign.xlHAlignRight },
        };

    private static readonly Dictionary<VerticalCellAlignment, Excel.XlVAlign> VerticalAlignment
        = new()
        {
            { VerticalCellAlignment.Bottom, Excel.XlVAlign.xlVAlignBottom },
            { VerticalCellAlignment.Center, Excel.XlVAlign.xlVAlignCenter },
            { VerticalCellAlignment.CenterAcrossSelection, Excel.XlVAlign.xlVAlignDistributed },
            { VerticalCellAlignment.Top, Excel.XlVAlign.xlVAlignTop },
        };

    public static void ClearFormatting()
    {
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        xlApp.ActiveWindow.DisplayGridlines = false;
        ((Excel.Worksheet)xlApp.ActiveSheet).DisplayPageBreaks = false;

        ((Excel.Range)xlApp.Selection).CurrentRegion.Select();
        var entireRange = (Excel.Range)xlApp.Selection;
        entireRange.Font.Name = "Calibri Light";
        entireRange.Font.Size = 10;
        SetRangeProperties(
            range: entireRange,
            bold: false,
            fontColor: rgbBlack,
            cellColor: rgbWhite,
            horizontalCellAlignment: HorizontalCellAlignment.Left,
            verticalCellAlignment: VerticalCellAlignment.Center);
        SetBorders(entireRange, false, true);
    }

    public static void FormatTable()
    {
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        xlApp.ActiveWindow.DisplayGridlines = false;
        ((Excel.Worksheet)xlApp.ActiveSheet).DisplayPageBreaks = false;

        ((Excel.Range)xlApp.Selection).CurrentRegion.Select();
        var entireRange = (Excel.Range)xlApp.Selection;
        entireRange.Font.Name = "Calibri Light";
        entireRange.Font.Size = 10;

        var tableOrientation = GetTableOrientation(entireRange);

        // What about a table that has both row and column labels?
        // var titleRange =
        //     tableOrientation == TableOrientation.Horizontal
        //         ? (Excel.Range)entireRange.Rows[1]
        //         : (Excel.Range)entireRange.Columns[1];
        var referenceTitle = GetReferenceTitleRow(entireRange);
        SetReferenceTitleRowFormatting(referenceTitle?.range);
        var primaryTitle = GetPrimaryTitleRow(entireRange, referenceTitle?.row ?? 0);
        SetPrimaryTitleRowFormatting(primaryTitle.range);
        var secondaryTitle = GetSecondaryTitleRow(entireRange, primaryTitle.row);
        SetSecondaryTitleRowFormatting(secondaryTitle?.range);
        var firstRowOrColumnOfBodyRange = secondaryTitle == null ? 2 : secondaryTitle?.row + 1;

        // The 1-based index of the first column containing content (i.e. content as opposed to titles) 
        int firstContentColumn = tableOrientation == TableOrientation.Vertical ? entireRange.Column : -1;

        var contentRange =
            tableOrientation == TableOrientation.Horizontal
                ? (Excel.Range)entireRange.Rows[$"{firstRowOrColumnOfBodyRange}:{entireRange.Rows.Count}"]
                : (Excel.Range)entireRange
                    .Columns[$"{GetColumnLetter(firstContentColumn)}:{GetColumnLetter(entireRange.Columns.Count)}"];

        SetBodyFormatting(contentRange, secondaryTitle?.range ?? primaryTitle.range, tableOrientation);
        entireRange.Columns.AutoFit();

        // format chart
        // if cell at the bottom has sum
    }

    private static string GetColumnLetter(int columnNumber)
    {
        var columnName = "";
        while (columnNumber > 0)
        {
            var modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }

        return columnName;
    }

    private static bool IsTitleRow(Excel.Range? row)
    {
        foreach (Excel.Range cell in row?.Cells!)
        {
            var cellValue = cell.Value2?.ToString() ?? "";
            if (double.TryParse(cellValue, out double x)
                || cellValue.Contains('<')
                || cellValue.Contains('>'))
            {
                return false;
            }
        }

        return true;
    }

    private static (Excel.Range? range, int row)? GetReferenceTitleRow(Excel.Range range)
    {

        foreach (Excel.Range cell in ((Excel.Range)range.Rows[1]).Cells)
        {
            var cellValue = cell.Value2?.ToString() ?? "";
            if (cellValue.Contains("<") || cellValue.Contains(">"))
            {
                return ((Excel.Range? range, int row))(range.Rows[1], 1);
            }
        }

        return null;
    }

    private static void SetReferenceTitleRowFormatting(Excel.Range? range)
    {
        if (range != null)
        {
            SetRangeProperties(
                range: range,
                bold: true,
                fontColor: rgbRed,
                cellColor: rgbWhite,
                horizontalCellAlignment: HorizontalCellAlignment.CenterAcrossSelection,
                verticalCellAlignment: VerticalCellAlignment.Center);
            SetBorders(range, false);
        }
    }

    private static (Excel.Range? range, int row) GetPrimaryTitleRow(Excel.Range range, int referenceTitleRow = 0)
    {
        // Check rows
        var currentRow = 0;
        for (int i = referenceTitleRow + 1; i <= 3; i++)
        {
            // Just check the first 2 rows.
            if (IsTitleRow(((Excel.Range)range.Rows[i]).Cells) || i >= 2)
            {
                currentRow = i;
                break;
            }
        }

        return ((Excel.Range)range.Rows[currentRow], currentRow);
        // Check columns
    }

    private static void SetPrimaryTitleRowFormatting(Excel.Range? range)
    {
        SetRangeProperties(
            range: range,
            bold: true,
            fontColor: rgbWhite,
            cellColor: rgbBlack,
            horizontalCellAlignment: HorizontalCellAlignment.CenterAcrossSelection,
            verticalCellAlignment: VerticalCellAlignment.Center);
        SetBorders(range, false);
    }

    private static (Excel.Range? range, int row)? GetSecondaryTitleRow(Excel.Range range, int primaryTitleRow)
    {
        if (IsTitleRow(range.Rows[primaryTitleRow + 1] as Excel.Range))
        {
            return ((Excel.Range)range.Rows[primaryTitleRow + 1], primaryTitleRow + 1);
        }
        else
        {
            return null;
        }
    }

    private static void SetSecondaryTitleRowFormatting(Excel.Range? range)
    {
        if (range != null)
        {
            SetRangeProperties(
                range: range,
                bold: true,
                fontColor: rgbWhite,
                cellColor: DeloitteGreen,
                horizontalCellAlignment: HorizontalCellAlignment.Center,
                verticalCellAlignment: VerticalCellAlignment.Center);
            SetBorders(range);
        }
    }

    private static (Excel.Range? range, int? row)? GetTotalsRow(Excel.Range? range)
    {
        var lastRow = range?.Rows.Count;
        foreach (Excel.Range cell in ((Excel.Range)range?.Rows[lastRow]!).Cells)
        {
            if (cell.Formula.ToString().ToUpper().Contains("SUM"))
            {
                return ((Excel.Range)range.Rows[lastRow]!, lastRow);
            }
        }

        return null;
    }

    private static void SetTotalsRowFormatting(Excel.Range? range)
    {
        if (range != null)
        {
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].Color = rgbBlack;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;

            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = rgbBlack;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;

            SetRangeProperties(
                range: range,
                bold: false,
                fontColor: rgbBlack,
                cellColor: rgbWhite,
                horizontalCellAlignment: HorizontalCellAlignment.CenterAcrossSelection,
                verticalCellAlignment: VerticalCellAlignment.Center);
        }
    }

    private static void SetBodyFormatting(Excel.Range range, Excel.Range? titleRange, TableOrientation tableOrientation)
    {
        range.Font.Bold = false;
        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        range.NumberFormat = "#,##0.00";

        foreach (var index in BoundaryBorderIndices)
        {
            range.Borders[index].Color = rgbBlack;
            range.Borders[index].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[index].Weight = Excel.XlBorderWeight.xlThin;
        }

        foreach (var index in InteriorBorderIndices)
        {
            range.Borders[index].Color = rgbBlack;
            range.Borders[index].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[index].Weight = Excel.XlBorderWeight.xlHairline;
        }

        var formats =
            new Dictionary<string[], string>
            {
                    { new[] { "(EUR" }, "[$€-x-euro2]#,##0.00" },
                    { new[] { "(GBP" }, "[$£-en-GB]#,##0.00" },
                    { new[] { "(RUB" }, "[$₽-ru-RU]#,##0.00" },
                    { new[] { "(USD" }, "[$$-en-US]#,##0.00" },
                    { new[] { "(ZAR" }, "[$R-en-ZA]#,##0.00" },
                    { new[]
                        {
                            "NACA",
                            "NACC",
                            "NACM",
                            "NACQ",
                            "NACS",
                            "Rate",
                            "Volatility",
                            "Vol",
                            "Yield"
                        }, "0.00%" },
                    { new[] { "Date" }, "yyyy-mm-dd;@" },
            };

        SetNumberFormattingBasedOnTitles(titleRange, range, formats, tableOrientation);
        var totalsRow = GetTotalsRow(range);
        SetTotalsRowFormatting(totalsRow?.range);
    }

    private static TableOrientation GetTableOrientation(Excel.Range range)
    {
        foreach (Excel.Range cell in ((Excel.Range)range.Columns[1]).Cells)
        {
            if (double.TryParse(cell.Value2.ToString(), out double _))
            {
                return TableOrientation.Horizontal;
            }
        }
        return TableOrientation.Vertical;
    }

    private static (bool ReferenceRangeExists, IEnumerable<(int rowNumber, int columnNumber)>)
        GetReferenceRange(Excel.Range range)
    {
        var referenceRangeExists = false;
        var referenceRange = new List<(int rowNumber, int columnNumber)>();
        foreach (Excel.Range cell in range.Cells)
        {
            if (cell.Value2.ToString().Contains("<") && cell.Value2.ToString().Contains(">"))
            {
                referenceRangeExists = true;
                referenceRange.Add((cell.Row, cell.Column));
            }
        }

        return (referenceRangeExists, referenceRange);
    }

    private static void SetBorders(Excel.Range? range, bool includeInteriorBorders = true, bool clearBorders = false)
    {
        foreach (var index in BoundaryBorderIndices)
        {
            range.Borders[index].Color = rgbBlack;
            range.Borders[index].Weight = Excel.XlBorderWeight.xlThin;
            range.Borders[index].LineStyle = clearBorders
                ? Excel.XlLineStyle.xlLineStyleNone
                : Excel.XlLineStyle.xlContinuous;
        }

        foreach (var index in InteriorBorderIndices)
        {
            range.Borders[index].Color = rgbBlack;
            range.Borders[index].Weight = Excel.XlBorderWeight.xlHairline;
            range.Borders[index].LineStyle = includeInteriorBorders && !clearBorders
                ? Excel.XlLineStyle.xlContinuous
                : Excel.XlLineStyle.xlLineStyleNone;
        }
    }

    private static void SetRangeProperties(
        Excel.Range? range,
        bool bold,
        object fontColor,
        object cellColor,
        HorizontalCellAlignment horizontalCellAlignment,
        VerticalCellAlignment verticalCellAlignment)
    {
        range.Font.Bold = bold;
        range.Font.Color = fontColor;
        range.HorizontalAlignment = HorizontalAlignment[horizontalCellAlignment];
        range.Interior.Color = cellColor;
        range.VerticalAlignment = VerticalAlignment[verticalCellAlignment];
    }

    // /// <summary>
    // /// Assumes '&lt;number&gt;' is in the first row.
    // /// </summary>
    // /// <param name="range"></param>
    // /// <returns></returns>
    // public Excel.Range GetReferenceTitleRange(Excel.Range range)
    // {
    //     var count = 0;
    //     foreach (Excel.Range cell in ((Excel.Range)range.Rows[1]).Cells)
    //     {
    //         var cellValue = cell.Value2.ToString();
    //         if (cellValue.Contains('<') && cellValue.Contains('>'))
    //         {
    //             count++;
    //         }
    //     }
    //
    //     if (count == 1)
    //     {
    //     }
    // }
    //

    private static void SetNumberFormattingBasedOnTitles(
        Excel.Range? titleRange,
        Excel.Range bodyRange,
        Dictionary<string[], string> formats,
        TableOrientation tableOrientation)
    {
        foreach (var format in formats)
        {
            var firstColumnIndex = ((Excel.Range)titleRange.Cells[1, 1]).Column;
            var firstRowIndex = ((Excel.Range)titleRange.Cells[1, 1]).Row;
            var indices = new List<int>();
            foreach (Excel.Range cell in titleRange.Cells)
            {
                foreach (var item in format.Key)
                {
                    if (cell.Value2.ToString().ToUpper().Contains(item.ToUpper()))
                    {
                        indices.Add(tableOrientation == TableOrientation.Horizontal
                            ? cell.Column - firstColumnIndex + 1
                            : cell.Row - firstRowIndex + 1);
                    }
                }
            }

            foreach (var index in indices)
            {
                if (tableOrientation == TableOrientation.Horizontal)
                {
                    ((Excel.Range)bodyRange.Columns[index]).NumberFormat = format.Value;
                }
                else
                {
                    ((Excel.Range)bodyRange.Rows[index]).NumberFormat = format.Value;
                }
            }
        }
    }
}
