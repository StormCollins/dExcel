using System.Text.RegularExpressions;
using dExcel.Dates;
using dExcel.Utilities;
using ExcelDna.Integration;
using QL = QuantLib;

namespace dExcel.ExcelUtils;

/// <summary>
/// A class for manipulating dExcel type tables in Excel.
/// </summary>
/// <remarks>Tables are assumed to be of the form:
/// Table Header
/// Column Header 1 | Column Header 2 | ... | Column Header n
/// Value 1         | Value 2         | ... | Value n
/// ...
/// </remarks>
public static class ExcelTableUtils
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
        List<string> columnHeaders
            = Enumerable
                .Range(0, table.GetLength(1))
                .Select(j => table[rowIndexOfColumnHeaders, j].ToString()?.ToUpper().Replace(" ", ""))
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
        int index = GetColumnHeaders(table, rowIndexOfColumnHeaders).IndexOf(columnHeader.ToUpper().Replace(" ", ""));
        if (index == -1)
        {
            return null;
        }

        if (typeof(T) == typeof(DateTime))
        {
            List<T> column =
                Enumerable
                    .Range(rowIndexOfColumnHeaders + 1, table.GetLength(0) - (rowIndexOfColumnHeaders + 1))
                    .Select(i => DateTime.FromOADate(int.Parse(table[i, index].ToString() ?? string.Empty)))
                    .Cast<T>()
                    .ToList();
            
            return column;
        }

        if (columnHeader.IgnoreCaseEquals("FRATenors"))
        {
            List<T> column =
                Enumerable
                    .Range(rowIndexOfColumnHeaders + 1, table.GetLength(0) - (rowIndexOfColumnHeaders + 1))
                    .Select(i => Regex.Match(table[i, index].ToString() ?? string.Empty, @"\d+(?=x)").Value)
                    .Select(startTenor => startTenor + "m")
                    .Cast<T>()
                    .ToList();
            
            return column;
        }
        else
        {
            List<T> column =
                Enumerable
                    .Range(rowIndexOfColumnHeaders + 1, table.GetLength(0) - (rowIndexOfColumnHeaders + 1))
                    .Select(i => (T)Convert.ChangeType(table[i, index], typeof(T)))
                    .ToList();
            
            return column;
        }
    }
    
    /// <summary>
    /// Gets a column from an Excel table given the zero-based column index.
    /// </summary>
    /// <param name="table">The input range.</param>
    /// <param name="columnIndex">The zero-based column index.</param>
    /// <typeparam name="T">The type to cast the column to e.g. "string" or "double".</typeparam>
    /// <returns>The table column.</returns>
    public static List<T> GetColumn<T>(object[,] table, int columnIndex = 0)
    {
        if (typeof(T) == typeof(DateTime))
        {
            List<T> column =
                Enumerable
                    .Range(0, table.GetLength(0))
                    .Select(i => DateTime.FromOADate(int.Parse(table[i, columnIndex].ToString() ?? string.Empty)))
                    .Cast<T>()
                    .ToList();
            
            return column;
        }
        else
        {
            List<T> column =
                Enumerable
                    .Range(0, table.GetLength(0))
                    .Select(i => (T)Convert.ChangeType(table[i, columnIndex], typeof(T)))
                    .ToList();
            
            return column;
        }
    }
    
    public static int GetRowIndex(object[,] table, string elementName, int columnIndex = 0)
    {
        List<string> column = GetColumn<string>(table, columnIndex);
        return column.IndexOf(elementName);
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
                .Select(i => table[i, 0].ToString()?.ToUpper().Replace(" ",  ""))
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
        int columnIndex = GetColumnHeaders(table, rowIndexOfColumnHeaders).IndexOf(columnHeader.ToUpper().Replace(" ", ""));
        if (columnIndex == -1)
        {
            return default;
        }
        
        int unadjustedRowIndex = GetRowHeaders(table, rowIndexOfColumnHeaders + 1).IndexOf(rowHeader.ToUpper().Replace(" ", ""));
        if (unadjustedRowIndex == -1)
        {
            return default;
        }
        
        int rowIndex = unadjustedRowIndex + rowIndexOfColumnHeaders + 1;

        if (typeof(T) == typeof(int))
        {
            return (T)Convert.ChangeType(
                value: int.Parse(table[rowIndex, columnIndex].ToString() ?? string.Empty), 
                conversionType: typeof(T));
        }

        if (typeof(T) == typeof(DateTime))
        {
            return (T)Convert.ChangeType(
                value: DateTime.FromOADate(int.Parse(table[rowIndex, columnIndex].ToString() ?? string.Empty)), 
                conversionType: typeof(T));
        }

        if (typeof(T) == typeof(QL.BusinessDayConvention))
        {
            (QL.BusinessDayConvention? businessDayConvention, string errorMessage) =
                DateUtils.ParseBusinessDayConvention(table[rowIndex, columnIndex].ToString() ?? string.Empty);

            if (businessDayConvention != null)
            {
                return (T)Convert.ChangeType(businessDayConvention, typeof(T));
            }

            throw new ArgumentException(
                CommonUtils.DExcelErrorMessage($"Invalid Business Day Convention: {table[rowIndex, columnIndex]}"));
        }

        if (typeof(T) == typeof(QL.DayCounter))
        {
            QL.DayCounter? dayCountConvention =
                DateUtils.ParseDayCountConvention(table[rowIndex, columnIndex].ToString() ?? string.Empty);
            
            if (dayCountConvention != null)
            {
                if (dayCountConvention.GetType() == typeof(QL.Business252))
                {
                    return (T)Convert.ChangeType(dayCountConvention, typeof(QL.Business252)); 
                }
                
                if (dayCountConvention.GetType() == typeof(QL.Actual360))
                {
                    return (T)Convert.ChangeType(dayCountConvention, typeof(QL.Actual360)); 
                }

                if (dayCountConvention.GetType() == typeof(QL.Actual365Fixed))
                {
                    return (T)Convert.ChangeType(dayCountConvention, typeof(QL.Actual365Fixed));
                }

                if (dayCountConvention.GetType() == typeof(QL.ActualActual))
                {
                    return (T)Convert.ChangeType(dayCountConvention, typeof(QL.ActualActual));
                }
            }

            throw new ArgumentException(
                CommonUtils.DExcelErrorMessage($"Invalid Day Count Convention: {table[rowIndex, columnIndex]}"));
        }
        
        return (T?)table[rowIndex, columnIndex];
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
        int columnIndex = GetColumnHeaders(table, rowIndexOfColumnHeaders).IndexOf(columnHeader.ToUpper());
        if (columnIndex == -1)
        {
            return null;
        }
        
        int rowIndexStart = GetRowHeaders(table, rowIndexOfColumnHeaders + 1).IndexOf(rowHeader.ToUpper());
        if (rowIndexStart == -1)
        {
            return null;
        }
        rowIndexStart += rowIndexOfColumnHeaders + 1;
        
        int rowIndexEnd = 1;
        
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
