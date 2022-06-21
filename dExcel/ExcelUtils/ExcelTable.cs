namespace dExcel.ExcelUtils;

using System.Collections.Generic;
using System.Linq;
using System.Windows.Documents;

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

    /// <summary>
    /// Gets a column from an Excel table given the column header.
    /// </summary>
    /// <param name="table">The input range.</param>
    /// <param name="columnHeader">The column header.</param>
    /// <typeparam name="T">The type to cast the column to e.g. "string" or "double".</typeparam>
    /// <returns>The table column.</returns>
    public static List<T> GetColumn<T>(object[,] table, string columnHeader)
    {
        var index = GetColumnHeaders(table).IndexOf(columnHeader);
        List<T> column;
        if (typeof(T) == typeof(double))
        {
            column =
                Enumerable
                    .Range(2, table.GetLength(0) - 2)
                    .Select(i => double.Parse(table[i, index].ToString()))
                    .Cast<T>()
                    .ToList();
        }
        else if (typeof(T) == typeof(int))
        {
            column =
                Enumerable
                    .Range(2, table.GetLength(0) - 2)
                    .Select(i => int.Parse(table[i, index].ToString()))
                    .Cast<T>()
                    .ToList();
        }
        else if (typeof(T) == typeof(DateTime))
        {
            column =
                Enumerable
                    .Range(2, table.GetLength(0) - 2)
                    .Select(i => DateTime.FromOADate(int.Parse(table[i, index].ToString())))
                    .Cast<T>()
                    .ToList();
        }
        else
        {
            column =
                Enumerable
                    .Range(2, table.GetLength(0) - 2)
                    .Select(i => table[i, index])
                    .Cast<T>()
                    .ToList();
        }
        
        return column;
    }
}
