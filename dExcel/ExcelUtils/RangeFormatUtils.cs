namespace dExcel.ExcelUtils;

using ExcelDna.Integration;
using static Microsoft.Office.Interop.Excel.XlRgbColor;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Used to specify the orientation of a table.
/// If it is 'Vertical' the data is organised into categories along columns e.g.,
///
/// | Date       | Cash Flow | Currency |
/// | 2022-06-30 | 1,000,000 |      ZAR |
/// | 2022-09-30 | 2,000,000 |      ZAR |
/// 
/// As opposed to 'Horizontal' where the data is organised into categories along rows e.g.,:
/// 
/// | Date       | 2022-06-30 | 2022-09-30 |
/// | Cash Flow  |  1,000,000 |  2,000,000 |
/// | Currency   |        ZAR |        ZAR |
/// 
/// If it is 'Both' then the data is usually primarily organised in categories along columns but the first column or two
/// also function as headers (this is common for vol surfaces) e.g.:
/// 
/// | Tenor     | 0.50 | 0.75 | 1.00 | 1.25 | 1.50 |
/// | 1w        |  10% |  12% |  15% |  18% |  20% |
/// | 1m        |  13% |  14% |  16% |  19% |  22% |
/// | 2m        |  14% |  16% |  18% |  21% |  24% |
/// | 3m        |  18% |  19% |  21% |  24% |  27% |
/// 
/// </summary>
public enum TableOrientation
{
    Both, Horizontal, Vertical, 
}

/// <summary>
/// Refers to how content is aligned horizontally within a cell.
/// </summary>
public enum HorizontalCellContentAlignment
{
    Center, CenterAcrossSelection, Left, Right,
}


/// <summary>
/// Refers to how content is aligned vertically within a cell.
/// </summary>
public enum VerticalCellContentAlignment
{
    Bottom, Center, CenterAcrossSelection, Top,
}


/// <summary>
/// A set of utilities for formatting Excel ranges.
/// </summary>
public static class RangeFormatUtils
{
    // Reviewed 
    // --------------------------------------------------------------------------------------
    private static readonly Color DeloitteGreen = Color.FromArgb(134, 188, 37);

    /// <summary>
    /// Checks if two ranges have the same formatting.
    /// </summary>
    /// <param name="cell1">Cell 1</param>
    /// <param name="cell2">Cell 2</param>
    /// <returns>True if Cell 1 and Cell 2 have the same formatting, otherwise false.</returns>
    [ExcelFunction(
        Name = "d.Formatting_Equal",
        Description = "Checks if the formatting of two cells are equal.",
        Category = "∂Excel: Formatting",
        IsMacroType = true)]
    public static bool Equal(
        [ExcelArgument(
            Name = "Cell 1",
            Description = "Cell 1",
            AllowReference = true)]
        object cell1,
        [ExcelArgument(
            Name = "Cell 2",
            Description = "Cell 2",
            AllowReference = true)]
        object cell2)
    {
        if (cell1 is ExcelReference reference1 && cell2 is ExcelReference reference2)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Range? range1 = (Excel.Range)xlApp.ActiveSheet.Cells[reference1.RowFirst + 1, reference1.ColumnFirst + 1];
            Excel.Range? range2 = (Excel.Range)xlApp.ActiveSheet.Cells[reference2.RowFirst + 1, reference2.ColumnFirst + 1];
            
            foreach (Excel.XlBordersIndex index in BoundaryBorderIndices)
            {
                if ((range1.Borders[index].Color != range2.Borders[index].Color) ||
                    (range1.Borders[index].LineStyle != range2.Borders[index].LineStyle) ||
                    (range1.Borders[index].Weight != range2.Borders[index].Weight))
                {
                    return false;
                }
            }

            return 
                range1.Font.Bold == range2.Font.Bold &&
                range1.Font.Color == range2.Font.Color &&
                range1.Font.Italic == range2.Font.Italic &&
                range1.Font.Name == range2.Font.Name &&
                range1.Font.Size == range2.Font.Size &&
                range1.Font.Underline == range2.Font.Underline &&
                range1.Interior.Color == range2.Interior.Color &&
                range1.MergeCells == range2.MergeCells &&
                range1.HorizontalAlignment == range2.HorizontalAlignment &&
                range1.VerticalAlignment == range2.VerticalAlignment;
        }
        
        return false;
    }

    /// <summary>
    /// Sets the formatting of a given range.
    /// </summary>
    /// <param name="range">The range to apply formatting to.</param>
    /// <param name="bold">True if the range font is bold.</param>
    /// <param name="fontColor">The font color.</param>
    /// <param name="cellBackgroundColor">The cell background color.</param>
    /// <param name="horizontalCellAlignment">The horizontal alignment setting of the cell contents.</param>
    /// <param name="verticalCellAlignment">The vertical alignment setting of the cell contents.</param>
    /// <param name="merge">Whether the cells should be merged. Default = False.</param>
    /// <param name="rotateText">Whether the contents of the cells should be rotated. Default = False.</param>
    private static void SetRangeFormatting(
        Excel.Range? range,
        bool bold,
        object fontColor,
        object cellBackgroundColor,
        HorizontalCellContentAlignment horizontalCellAlignment,
        VerticalCellContentAlignment verticalCellAlignment,
        bool merge = false,
        bool rotateText = false)
    {
        if (range != null)
        {
            range.Font.Bold = bold;
            range.Font.Color = fontColor;
            range.HorizontalAlignment = HorizontalAlignment[horizontalCellAlignment];
            range.Interior.Color = cellBackgroundColor;
            range.VerticalAlignment = VerticalAlignment[verticalCellAlignment];
            range.MergeCells = merge;
            range.Orientation = rotateText ? 90 : 0;
        }
    }

    /// <summary>
    /// Removes the following formatting from a contiguous region:
    /// <list type="bullet">
    ///     <item>
    ///         <description>Background color</description>
    ///     </item>
    ///     <item>
    ///         <description>Borders</description>
    ///     </item>
    /// </list>
    /// And sets the font to the following:
    /// <list type="bullet">
    ///     <item>
    ///         <description>Calibri light</description>
    ///     </item>
    ///     <item>
    ///         <description>Black</description>
    ///     </item>
    ///     <item>
    ///         <description>Size 10</description>
    ///     </item>
    ///     <item>
    ///         <description>No boldness</description>
    ///     </item>
    ///     <item>
    ///         <description>No italics</description>
    ///     </item>
    ///     <item>
    ///         <description>No underlining</description>
    ///    </item>
    /// </list>
    /// </summary>
    [ExcelFunction(
        Name = "d.Formatting_ClearRangeFormatting",
        Description = "Clears the formatting of a table.",
        Category = "∂Excel: Formatting")]
    public static void ClearRangeFormatting()
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        xlApp.ActiveWindow.DisplayGridlines = false;
        ((Excel.Worksheet)xlApp.ActiveSheet).DisplayPageBreaks = false;

        ((Excel.Range)xlApp.Selection).CurrentRegion.Select();
        Excel.Range? entireRange = (Excel.Range)xlApp.Selection;
        entireRange.Font.Name = "Calibri Light";
        entireRange.Font.Size = 10;
        SetRangeFormatting(
            range: entireRange,
            bold: false,
            fontColor: rgbBlack,
            cellBackgroundColor: rgbWhite,
            horizontalCellAlignment: HorizontalCellContentAlignment.Left,
            verticalCellAlignment: VerticalCellContentAlignment.Center);
        entireRange.NumberFormat = "#,##0.00";
        SetBorders(entireRange, false, true);
    }

    /// <summary>
    /// Sets the standard formatting for a primary title row or column i.e. a white font on black background.
    /// </summary>
    /// <param name="range">The range to apply the formatting to.</param>
    /// <param name="tableOrientation">Whether the table is 'vertically' or 'horizontally' aligned i.e., has column or
    /// row headers. 'Both' is not applicable here.</param>
    /// <param name="merge">Whether the cells should be merged. Default = true.</param>
    /// <param name="rotateText">Whether the contents of the cells should be rotated. Default = true.</param>
    public static void SetPrimaryTitleRangeFormatting(
        Excel.Range? range, 
        TableOrientation tableOrientation = TableOrientation.Vertical, 
        bool merge = true, 
        bool rotateText = true)
    {
        if (tableOrientation == TableOrientation.Vertical)
        {
            SetRangeFormatting(
                range: range,
                bold: true,
                fontColor: rgbWhite,
                cellBackgroundColor: rgbBlack,
                horizontalCellAlignment: HorizontalCellContentAlignment.CenterAcrossSelection,
                verticalCellAlignment: VerticalCellContentAlignment.Center);
        }
        else
        {
            SetRangeFormatting(
                range: range,
                bold: true,
                fontColor: rgbWhite,
                cellBackgroundColor: rgbBlack,
                horizontalCellAlignment: HorizontalCellContentAlignment.CenterAcrossSelection,
                verticalCellAlignment: VerticalCellContentAlignment.Center,
                merge,
                rotateText);
        }
        
        SetBorders(range, false);
    }

    /// <summary>
    /// Sets the standard formatting for a secondary title row or column i.e. a white font on a green background.
    /// </summary>
    /// <param name="range">The range to apply the formatting to.</param>
    public static void SetSecondaryTitleRangeFormatting(Excel.Range? range)
    {
        if (range != null)
        {
            SetRangeFormatting(
                range: range,
                bold: true,
                fontColor: rgbWhite,
                cellBackgroundColor: DeloitteGreen,
                horizontalCellAlignment: HorizontalCellContentAlignment.Center,
                verticalCellAlignment: VerticalCellContentAlignment.Center);
            SetBorders(range);
        }
    }

    /// <summary>
    /// Sets standard sheet wide formatting.
    /// </summary>
    public static void SetSheetWideFormatting()
    {
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        xlApp.ActiveWindow.DisplayGridlines = false;
        ((Excel.Worksheet)xlApp.ActiveSheet).DisplayPageBreaks = false;
    }

    /// <summary>
    /// Sets the formatting for a table which is vertically aligned i.e., has column headers with data 
    /// running column-wise.
    /// </summary>
    /// <param name="hasTwoHeaderRows">True if the table has two header rows.</param>
    [ExcelFunction(
        Name = "d.Formatting_SetVerticallyAlignedTableFormatting",
        Description = "Sets the default formatting for a vertically aligned (i.e. column-based) table.",
        Category = "∂Excel: Formatting")]
    public static void SetVerticallyAlignedTableFormatting(bool hasTwoHeaderRows)
    {
        SetSheetWideFormatting();
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        ((Excel.Range)xlApp.Selection).CurrentRegion.Select();
        var entireRange = (Excel.Range)xlApp.Selection;
        entireRange.Font.Name = "Calibri Light";
        entireRange.Font.Size = 10;

        SetPrimaryTitleRangeFormatting(entireRange.Rows[1]);

        if (hasTwoHeaderRows)
        {
            var columnHeaders = entireRange.Rows[2];
            SetSecondaryTitleRangeFormatting(columnHeaders);
            SetTableContentFormatting(entireRange.Rows[$"3:{entireRange.Rows.Count}"], columnHeaders, TableOrientation.Vertical);
        }
        else
        {
            var columnHeaders = entireRange.Rows[1];
            SetTableContentFormatting(entireRange.Rows[$"2:{entireRange.Rows.Count}"], columnHeaders, TableOrientation.Vertical);
        }
    }

    /// <summary>
    /// Sets the formatting for a table which is horizontally aligned i.e. has row headers with data 
    /// running row-wise.
    /// </summary>
    /// <param name="hasTwoHeaderColumns">True if the table has two header columns.</param>
    [ExcelFunction(
        Name = "d.Formatting_SetHorizontallyAlignedTableFormatting",
        Description = "Sets the default formatting for a horizontally aligned (i.e., row-based) table.",
        Category = "∂Excel: Formatting")]
    public static void SetHorizontallyAlignedTableFormatting(bool hasTwoHeaderColumns)
    {
        SetSheetWideFormatting();
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        ((Excel.Range)xlApp.Selection).CurrentRegion.Select();
        var entireRange = (Excel.Range)xlApp.Selection;
        entireRange.Font.Name = "Calibri Light";
        entireRange.Font.Size = 10;


        if (hasTwoHeaderColumns)
        {
            SetPrimaryTitleRangeFormatting(entireRange.Columns[1], TableOrientation.Horizontal);
            var rowHeaders = entireRange.Columns[2];
            SetSecondaryTitleRangeFormatting(rowHeaders);
            SetTableContentFormatting(entireRange.Columns[$"{GetColumnLetter(3)}:{GetColumnLetter(entireRange.Columns.Count)}"], rowHeaders, TableOrientation.Horizontal);
        }
        else
        {
            SetPrimaryTitleRangeFormatting(entireRange.Columns[1], TableOrientation.Horizontal, false, false);
            var rowHeaders = entireRange.Columns[1];
            SetTableContentFormatting(entireRange.Columns[$"{GetColumnLetter(2)}:{GetColumnLetter(entireRange.Columns.Count)}"], rowHeaders, TableOrientation.Horizontal);
        }
    }

    /// <summary>
    /// Sets the formatting for a table which is horizontally aligned i.e. has row headers with data 
    /// running row-wise.
    /// </summary>
    /// <param name="hasTwoHeaderRows">True if the table has two header rows.</param>
    /// <param name="hasTwoHeaderColumns">True if the table has two header columns.</param>
    [ExcelFunction(
        Name = "d.Formatting_SetHorizontalAndVerticalTableFormatting",
        Description = "Sets the default formatting a table which has both column and row headers.",
        Category = "∂Excel: Formatting")]
    public static void SetColumnAndRowHeaderBasedTableFormatting(bool hasTwoHeaderRows, bool hasTwoHeaderColumns)
    {
        SetSheetWideFormatting();
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        ((Excel.Range)xlApp.Selection).CurrentRegion.Select();
        var entireRange = (Excel.Range)xlApp.Selection;
        entireRange.Font.Name = "Calibri Light";
        entireRange.Font.Size = 10;

        Excel.Range primaryRow = entireRange.Rows[1];
        Excel.Range secondaryRow = entireRange.Rows[2];
        Excel.Range primaryColumnHeaders;

        if (hasTwoHeaderRows)
        {
            primaryColumnHeaders = primaryRow.Columns[$"{GetColumnLetter(3)}:{GetColumnLetter(entireRange.Columns.Count)}"];
            SetPrimaryTitleRangeFormatting(primaryColumnHeaders);

            Excel.Range secondaryColumnHeaders;
            if (hasTwoHeaderColumns)
            {
                secondaryColumnHeaders =
                    secondaryRow.Columns[$"{GetColumnLetter(3)}:{GetColumnLetter(entireRange.Columns.Count)}"];
                SetTableContentFormatting(entireRange.Rows[$"3:{entireRange.Rows.Count}"].Columns[$"{GetColumnLetter(3)}:{GetColumnLetter(entireRange.Columns.Count)}"], secondaryColumnHeaders, TableOrientation.Vertical);
            }
            else
            {
                secondaryColumnHeaders =
                    secondaryRow.Columns[$"{GetColumnLetter(2)}:{GetColumnLetter(entireRange.Columns.Count)}"];
                SetTableContentFormatting(entireRange.Rows[$"3:{entireRange.Rows.Count}"].Columns[$"{GetColumnLetter(2)}:{GetColumnLetter(entireRange.Columns.Count)}"], secondaryColumnHeaders, TableOrientation.Vertical);
            }

            SetSecondaryTitleRangeFormatting(secondaryColumnHeaders);
        }
        else
        {
            primaryColumnHeaders = primaryRow.Columns[$"{GetColumnLetter(2)}:{GetColumnLetter(entireRange.Columns.Count)}"];
            SetPrimaryTitleRangeFormatting(primaryColumnHeaders);
        }

        Excel.Range primaryColumn = entireRange.Columns[1];
        Excel.Range secondaryColumn = entireRange.Columns[2];
        Excel.Range primaryRowHeaders;

        if (hasTwoHeaderColumns)
        {
            primaryRowHeaders = primaryColumn.Rows[$"3:{entireRange.Rows.Count}"];
            Excel.Range secondaryRowHeaders;

            if (hasTwoHeaderRows)
            {
                secondaryRowHeaders = secondaryColumn.Rows[$"3:{entireRange.Rows.Count}"];
                SetPrimaryTitleRangeFormatting(primaryRowHeaders, TableOrientation.Horizontal);
            }
            else
            {
                secondaryRowHeaders = secondaryColumn.Rows[$"2:{entireRange.Rows.Count}"];
                SetPrimaryTitleRangeFormatting(primaryRowHeaders, TableOrientation.Horizontal, false, false);
            }

            SetSecondaryTitleRangeFormatting(secondaryRowHeaders);
        }
        else
        {
            primaryRowHeaders = primaryColumn.Rows[$"2:{entireRange.Rows.Count}"];
            SetPrimaryTitleRangeFormatting(primaryRowHeaders, TableOrientation.Horizontal, false, false);
        }
    }

    /// <summary>
    /// Sets the formatting of the table content i.e. not the header rows nor columns.
    /// </summary>
    /// <param name="contentRange">The content range.</param>
    /// <param name="headerRange">The header range from which to infer the numeric and date formatting.</param>
    /// <param name="tableOrientation">Specifies whether table has column or row headers.</param>
    private static void SetTableContentFormatting(
        Excel.Range contentRange,
        Excel.Range headerRange,
        TableOrientation tableOrientation)
    {
        contentRange.Font.Bold = false;
        contentRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        contentRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        contentRange.NumberFormat = "#,##0.00";

        // Set out border format.
        foreach (var index in BoundaryBorderIndices)
        {
            contentRange.Borders[index].Color = rgbBlack;
            contentRange.Borders[index].LineStyle = Excel.XlLineStyle.xlContinuous;
            contentRange.Borders[index].Weight = Excel.XlBorderWeight.xlThin;
        }

        // Set inner border format.
        foreach (var index in InteriorBorderIndices)
        {
            contentRange.Borders[index].Color = rgbBlack;
            contentRange.Borders[index].LineStyle = Excel.XlLineStyle.xlContinuous;
            contentRange.Borders[index].Weight = Excel.XlBorderWeight.xlHairline;
        }

        var numericalAndDateFormats =
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

        SetNumericalAndDateFormattingBasedOnHeaders(headerRange, contentRange, numericalAndDateFormats, tableOrientation);
        var totalsRow = GetTotalsRow(contentRange);
        SetTotalsRowFormatting(totalsRow?.range);
    }





    // Not Reviewed 
    // --------------------------------------------------------------------------------------

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

    private static readonly Dictionary<HorizontalCellContentAlignment, Excel.XlHAlign> HorizontalAlignment
        = new()
        {
            { HorizontalCellContentAlignment.Center, Excel.XlHAlign.xlHAlignCenter },
            { HorizontalCellContentAlignment.CenterAcrossSelection, Excel.XlHAlign.xlHAlignCenterAcrossSelection },
            { HorizontalCellContentAlignment.Left, Excel.XlHAlign.xlHAlignLeft },
            { HorizontalCellContentAlignment.Right, Excel.XlHAlign.xlHAlignRight },
        };

    private static readonly Dictionary<VerticalCellContentAlignment, Excel.XlVAlign> VerticalAlignment
        = new()
        {
            { VerticalCellContentAlignment.Bottom, Excel.XlVAlign.xlVAlignBottom },
            { VerticalCellContentAlignment.Center, Excel.XlVAlign.xlVAlignCenter },
            { VerticalCellContentAlignment.CenterAcrossSelection, Excel.XlVAlign.xlVAlignDistributed },
            { VerticalCellContentAlignment.Top, Excel.XlVAlign.xlVAlignTop },
        };



    /// <summary>
    /// Gets the equivalent column letter for a column number e.g. A = 1, B = 2, etc.
    /// </summary>
    /// <param name="columnNumber">The column number.</param>
    /// <returns>The column letter.</returns>
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

            SetRangeFormatting(
                range: range,
                bold: false,
                fontColor: rgbBlack,
                cellBackgroundColor: rgbWhite,
                horizontalCellAlignment: HorizontalCellContentAlignment.CenterAcrossSelection,
                verticalCellAlignment: VerticalCellContentAlignment.Center);
        }
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


    private static void SetNumericalAndDateFormattingBasedOnHeaders(
        Excel.Range? headerRange,
        Excel.Range contentRange,
        Dictionary<string[], string> formats,
        TableOrientation tableOrientation)
    {
        foreach (var format in formats)
        {
            var firstColumnIndex = ((Excel.Range)headerRange.Cells[1, 1]).Column;
            var firstRowIndex = ((Excel.Range)headerRange.Cells[1, 1]).Row;
            var indices = new List<int>();
            foreach (Excel.Range cell in headerRange.Cells)
            {
                foreach (var item in format.Key)
                {
                    if (cell.Value2.ToString().ToUpper().Contains(item.ToUpper()))
                    {
                        indices.Add(tableOrientation == TableOrientation.Vertical
                            ? cell.Column - firstColumnIndex + 1
                            : cell.Row - firstRowIndex + 1);
                    }
                }
            }

            foreach (var index in indices)
            {
                if (tableOrientation == TableOrientation.Vertical)
                {
                    ((Excel.Range)contentRange.Columns[index]).NumberFormat = format.Value;
                    ((Excel.Range)contentRange.Columns[index]).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                }
                else
                {
                    ((Excel.Range)contentRange.Rows[index]).NumberFormat = format.Value;
                    ((Excel.Range)contentRange.Rows[index]).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                }
            }
        }
    }

    public static void SetConditionalTestUtilsFormatting()
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Range? selectedRange = (Excel.Range)xlApp.Selection;
        selectedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        selectedRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

        Excel.FormatCondition errorFormatCondition = (Excel.FormatCondition)selectedRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "=\"ERROR\"");
        errorFormatCondition.Font.Bold = true;
        errorFormatCondition.Font.Color = -16383844;
        errorFormatCondition.Font.TintAndShade = 0;
        errorFormatCondition.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
        errorFormatCondition.Interior.Color = 13551615;
        errorFormatCondition.Interior.TintAndShade = 0;
        errorFormatCondition.Borders.LineStyle = Excel.XlLineStyle.xlDash;
        errorFormatCondition.Borders.Color = -16383844;
        errorFormatCondition.Borders.TintAndShade = 0;
        errorFormatCondition.Borders.Weight = Excel.XlBorderWeight.xlThin;

        Excel.FormatCondition okFormatCondition = (Excel.FormatCondition)selectedRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "=\"OK\"");
        okFormatCondition.Font.Bold = true;
        okFormatCondition.Font.Color = -16752384;
        okFormatCondition.Font.TintAndShade = 0;
        okFormatCondition.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
        okFormatCondition.Interior.Color = 13561798;
        okFormatCondition.Interior.TintAndShade = 0;
        okFormatCondition.Borders.LineStyle = Excel.XlLineStyle.xlDash;
        okFormatCondition.Borders.Color = -16752384;
        okFormatCondition.Borders.TintAndShade = 0;
        okFormatCondition.Borders.Weight = Excel.XlBorderWeight.xlThin;

        Excel.FormatCondition warningFormatCondition = (Excel.FormatCondition)selectedRange.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "=\"WARNING\"");
        warningFormatCondition.Font.Bold = true;
        warningFormatCondition.Font.Color = -16754788;
        warningFormatCondition.Font.TintAndShade = 0;
        warningFormatCondition.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
        warningFormatCondition.Interior.Color = 10284031;
        warningFormatCondition.Interior.TintAndShade = 0;
        warningFormatCondition.Borders.LineStyle = Excel.XlLineStyle.xlDash;
        warningFormatCondition.Borders.Color = -16754788;
        warningFormatCondition.Borders.TintAndShade = 0;
        warningFormatCondition.Borders.Weight = Excel.XlBorderWeight.xlThin;
    }

}
