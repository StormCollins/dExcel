using ExcelDna.Integration;

namespace dExcel.ExcelUtils;

/// <summary>
/// A collection of utilities for dealing with worksheets in Excel.
/// </summary>
public static class SheetUtils
{
    /// <summary>
    /// Gets the current sheet name.
    /// </summary>
    /// <returns>The current sheet name.</returns>
    [ExcelFunction(
        Name = "d.Excel_GetSheetName",
        Description = "Gets the current sheet name.",
        Category = "∂Excel: Excel Utils")]
    public static string GetSheetName(bool removeSpaces = true)
    {
        ExcelReference reference = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
        string sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
        sheetName = sheetName[(sheetName.LastIndexOf(']') + 1)..];
        if (removeSpaces)
        {
            sheetName = sheetName.Replace(" ", "");
        }

        return sheetName;
    }
}
