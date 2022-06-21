namespace dExcel.ExcelUtils;

using System.Collections.Generic;
using System.Linq;

/// <summary>
/// A class for manipulating dExcel type tables in Excel.
/// </summary>
/// <remarks>Tables are assumed to be of the form:
/// Table Type
/// Column Header 1 | Column Header 2 | ... | Column Header n
/// Value 1         | Value 2         | ... | Value n
/// ...
/// </remarks>
public static class ExcelTable
{
    /// <summary>
    /// Gets the table type.
    /// </summary>
    /// <remarks>It is assumed the table type is at position [0,0] in the 2D array.</remarks>
    /// <param name="table">The input range.</param>
    /// <returns>The table type.</returns>
    public static string? GetTableType(object[,] table)
    {
        return table[0, 0].ToString();
    }
    
    /// <summary>
    /// Get the list of column headers of a table.
    /// </summary>
    /// <remarks>Assumes the column headers are in row 1 (of a zero based row index).</remarks>
    /// <param name="table">The input range.</param>
    /// <returns>The list of column headers.</returns>
    public static List<string> GetColumnHeaders(object[,] table)
    {
        List<string> columnTitles
                = Enumerable
                    .Range(0, table.GetLength(1))
                    .Select(j => table[1, j])
                    .Cast<string>()
                    .ToList();
        return columnTitles;
    }
}
