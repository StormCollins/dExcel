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
        var columnHeaders
            = Enumerable
                .Range(0, table.GetLength(1))
                .Select(j => table[1, j])
                .Cast<string>()
                .ToList();
        return columnHeaders;
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
        if (typeof(T) == typeof(DateTime))
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
                    .Select(i => (T)Convert.ChangeType(table[i, index], typeof(T)))
                    .ToList();
        }
        return column;
    }
    
    /// <summary>
    /// Get the list of row headers of a table.
    /// </summary>
    /// <remarks>Assumes the row headers are in column 0 and start from row 2.</remarks>
    /// <param name="table">The input range.</param>
    /// <returns>The list of row headers.</returns>
    public static List<string> GetRowHeaders(object[,] table)
    {
        var rowHeaders
            = Enumerable
                .Range(2, table.GetLength(0) - 2)
                .Select(i => table[i, 0])
                .Cast<string>()
                .ToList(); 
        return rowHeaders;
    }

    /// <summary>
    /// Looks up a value in an Excel table using a column and row header. Assumes row headers are in column 0 and column
    /// headers are in row 2.
    /// </summary>
    /// <param name="table">The Excel input range.</param>
    /// <param name="columnHeader">The column header.</param>
    /// <param name="rowHeader">The row header.</param>
    /// <returns>The looked up value.</returns>
    public static T LookupTableValue<T>(object[,] table, string columnHeader, string rowHeader)
    {
        var columnIndex = GetColumnHeaders(table).IndexOf(columnHeader);
        var rowIndex = GetRowHeaders(table).IndexOf(rowHeader) + 2;
        if (typeof(T) == typeof(DateTime))
        {
            return (T)Convert.ChangeType(DateTime.FromOADate(int.Parse(table[rowIndex, columnIndex].ToString())), typeof(T));
        }
        return (T)Convert.ChangeType(table[rowIndex, columnIndex], typeof(T));
    }
}
