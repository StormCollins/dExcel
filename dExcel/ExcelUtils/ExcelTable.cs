namespace dExcel.ExcelUtils;

using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using QLNet;

/// <summary>
/// A class for manipulating dExcel type tables in Excel.
/// </summary>
/// <remarks>Tables are assumed to be of the form:
/// Table Header
/// Column Header 1 | Column Header 2 | ... | Column Header n
/// Value 1         | Value 2         | ... | Value n
/// ...
/// </remarks>
public static class ExcelTable
{
    /// <summary>
    /// Gets the table label.
    /// </summary>
    /// <remarks>It is assumed the table label is at position [0,0] in the 2D array.</remarks>
    /// <param name="table">The input range.</param>
    /// <returns>The table label.</returns>
    public static string? GetTableLabel(object[,] table)
    {
        return table[0, 0].ToString();
    }
    
    /// <summary>
    /// Get the list of column headers of a table.
    /// </summary>
    /// <remarks>Since tables typically have a table header, we assume the column headers are in row 1 (of a zero based
    /// row index).</remarks>
    /// <param name="table">The input range.</param>
    /// <param name="rowIndexOfColumnHeaders">The index of the row containing the column headers.</param>
    /// <returns>The list of column headers.</returns>
    public static List<string> GetColumnHeaders(object[,] table, int rowIndexOfColumnHeaders = 1)
    {
        var columnHeaders
            = Enumerable
                .Range(0, table.GetLength(1))
                .Select(j => table[rowIndexOfColumnHeaders, j])
                .Cast<string>()
                .ToList();
        return columnHeaders;
    }

    /// <summary>
    /// Gets a column from an Excel table given the column header name.
    /// </summary>
    /// <remarks>Since tables typically have a table header, we assume the column headers are in row 1 (of a zero based
    /// row index).</remarks>
    /// <param name="table">The input range.</param>
    /// <param name="columnHeader">The column header.</param>
    /// <param name="rowIndexOfColumnHeaders">The index of the row containing the column headers.</param>
    /// <typeparam name="T">The type to cast the column to e.g. "string" or "double".</typeparam>
    /// <returns>The table column.</returns>
    public static List<T>? GetColumn<T>(object[,] table, string columnHeader, int rowIndexOfColumnHeaders = 1)
    {
        var index = GetColumnHeaders(table, rowIndexOfColumnHeaders).IndexOf(columnHeader);
        if (index == -1)
        {
            return null;
        }

        if (typeof(T) == typeof(DateTime))
        {
            var column =
                Enumerable
                    .Range(rowIndexOfColumnHeaders + 1, table.GetLength(0) - (rowIndexOfColumnHeaders + 1))
                    .Select(i => DateTime.FromOADate(int.Parse(table[i, index].ToString())))
                    .Cast<T>()
                    .ToList();
            return column;
        }
        else if (string.Compare(columnHeader, "FRATenors", StringComparison.InvariantCultureIgnoreCase) == 0)
        {
            var column =
                Enumerable
                    .Range(rowIndexOfColumnHeaders + 1, table.GetLength(0) - (rowIndexOfColumnHeaders + 1))
                    .Select(i => Regex.Match(table[i, index].ToString(), @"\d+(?=x)").Value)
                    .Select(startTenor => startTenor + "m")
                    .Cast<T>()
                    .ToList();
            return column;
        }
        else
        {
            var column =
                Enumerable
                    .Range(rowIndexOfColumnHeaders + 1, table.GetLength(0) - (rowIndexOfColumnHeaders + 1))
                    .Select(i => (T)Convert.ChangeType(table[i, index], typeof(T)))
                    .ToList();
            return column;
        }
        
    }

    /// <summary>
    /// Gets the list of row headers of a table.
    /// </summary>
    /// <remarks>Since tables typically have a table header, we assume the column headers are in row 1 (of a zero based
    /// row index) and thus the row headers start at row 2.</remarks>
    /// <param name="table">The input range.</param>
    /// <param name="rowIndexOfFirstRowHeader">The row index of the first row header.</param>
    /// <returns>The list of row headers.</returns>
    public static List<string?> GetRowHeaders(object[,] table, int rowIndexOfFirstRowHeader = 2)
    {
        return Enumerable
                .Range(rowIndexOfFirstRowHeader, table.GetLength(0) - rowIndexOfFirstRowHeader)
                .Select(i => table[i, 0].ToString())
                .ToList(); 
    }

    /// <summary>
    /// Looks up a value in an Excel table using a column and row header. Assumes row headers are in column 0 and column
    /// headers are in row 2.
    /// </summary>
    /// <remarks>Since tables typically have a table header, we assume the column headers are in row 1 (of a zero based
    /// row index).</remarks>
    /// <param name="table">The Excel input range.</param>
    /// <param name="columnHeader">The column header.</param>
    /// <param name="rowHeader">The row header.</param>
    /// <param name="rowIndexOfColumnHeaders">The index of the row containing the column headers.</param>
    /// <returns>The looked up value.</returns>
    public static T? GetTableValue<T>(
        object[,] table,
        string columnHeader,
        string rowHeader,
        int rowIndexOfColumnHeaders = 1)
    {
        var columnIndex = GetColumnHeaders(table, rowIndexOfColumnHeaders).IndexOf(columnHeader);
        var rowIndex = GetRowHeaders(table, rowIndexOfColumnHeaders + 1).IndexOf(rowHeader) + rowIndexOfColumnHeaders + 1;
        if (columnIndex == -1  || rowIndex <= rowIndexOfColumnHeaders)
        {
            return default;
        }

        if (typeof(T) == typeof(DateTime))
        {
            return (T)Convert.ChangeType(DateTime.FromOADate(int.Parse(table[rowIndex, columnIndex].ToString())), typeof(T));
        }
        else if (typeof(T) == typeof(BusinessDayConvention))
        {
            BusinessDayConvention? businessDayConvention =
                table[rowIndex, columnIndex]?.ToString()?.ToUpper() switch
                {
                    "FOLLOWING" or "FOL" => BusinessDayConvention.Following,
                    "MODIFIEDFOLLOWING" or "MODFOL" => BusinessDayConvention.ModifiedFollowing,
                    "MODIFIEDPRECEDING" or "MODPREC" => BusinessDayConvention.ModifiedPreceding,
                    "PRECEDING" or "PREC" => BusinessDayConvention.Preceding,
                    _ => null,
                };

            if (businessDayConvention != null)
            {
                return (T)Convert.ChangeType(businessDayConvention, typeof(T));
            }
            else
            {
                return default;
            }
        }
        else if (typeof(T) == typeof(DayCounter))
        {
            DayCounter? dayCountConvetion =
                table[rowIndex, columnIndex]?.ToString()?.ToUpper() switch
                {
                    "ACTUAL360" or "ACT360" => new Actual360(),
                    "ACTUAL365" or "ACT365" => new Actual365Fixed(),
                    "ACTUALACTUAL" or "ACTACT" => new ActualActual(),
                    "BUSINESS252" => new Business252(),
                    _ => null,
                };
            if (dayCountConvetion != null)
            {
                return (T)Convert.ChangeType(dayCountConvetion, typeof(T));
            }
            else
            {
                return default;
            }
        }
        else
        {
            return (T)Convert.ChangeType(table[rowIndex, columnIndex], typeof(T));
        }
    }
    
    /// <summary>
    /// Looks up several values in an Excel table using a column and row header where multiple values fall under the
    /// same row header. Assumes row headers are in column 0 and column headers are in row 2.
    /// </summary>
    /// <remarks>If the table is as follows:
    /// Table Type
    /// Parameter     | Value 
    /// Instruments   | Deposits
    ///               | FRAs
    ///               | Interest Rate Swaps
    /// Interpolation | LogLinear
    /// ...
    /// Using this function to lookup the "Instruments" it will return: 'Deposits', 'FRAs', and 'Interest Rate Swaps'.
    /// 
    /// Since tables typically have a table header, we assume the column headers are in row 1 (of a zero based
    /// row index).
    /// </remarks>
    /// <param name="table">The Excel input range.</param>
    /// <param name="columnHeader">The column header.</param>
    /// <param name="rowHeader">The row header.</param>
    /// <param name="rowIndexOfColumnHeaders">The index of the row containing the column headers.</param>
    /// <returns>The looked up value. If it can't find the column/row header, returns null.</returns>
    public static List<T>? LookUpTableValues<T>(
        object[,] table,
        string columnHeader,
        string rowHeader,
        int rowIndexOfColumnHeaders = 1)
    {
        var columnIndex = GetColumnHeaders(table, rowIndexOfColumnHeaders).IndexOf(columnHeader);
        if (columnIndex == -1)
        {
            return null;
        }
        
        var rowIndexStart = GetRowHeaders(table, rowIndexOfColumnHeaders + 1).IndexOf(rowHeader);
        if (rowIndexStart == -1)
        {
            return null;
        }
        rowIndexStart += rowIndexOfColumnHeaders + 1;
        
        var rowIndexEnd = 1;
        
        for (int i = rowIndexStart + 1; i < table.GetLength(0); i++)
        {
            if (table[i, 0].ToString() == "" || table[i, 0] == ExcelMissing.Value)
            {
                rowIndexEnd++;
            }
            else
            {
                break;
            }
        }

        return Enumerable
                .Range(rowIndexStart, rowIndexEnd)
                .Select(i => (T)Convert.ChangeType(table[i, columnIndex], typeof(T)))
                .ToList();
    }
}
