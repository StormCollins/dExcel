namespace dExcel.ExcelUtils;

/// <summary>
/// This class provides a set of utility methods for between Excel ranges and native enumerable, C# types.
/// </summary>
public static class ExcelArrayUtils
{
    /// <summary>
    /// Converts an Excel range, either a single column or row, it does not handle 2D ranges, to a list.
    /// </summary>
    /// <param name="range">The Excel range to convert.</param>
    /// <typeparam name="T">The type to convert the range to.</typeparam>
    /// <returns>A list of values contained in the Excel range.</returns>
    public static List<T> ConvertExcelRangeToList<T>(object[,] range)
    {
        List<T> output = new();

        int dimension = range.GetLength(0) > range.GetLength(1) ? 0 : 1;
        if (typeof(T) == typeof(DateTime))
        {
            for (int i = 0; i < range.GetLength(dimension); i++)
            {
                if (dimension == 0)
                {
                    output.Add((T)Convert.ChangeType(DateTime.FromOADate((double)range[i, 0]), typeof(T)));
                }
                else
                {
                    output.Add((T)Convert.ChangeType(DateTime.FromOADate((double)range[0, i]), typeof(T)));
                }
            }
        }
        else
        {
            for (int i = 0; i < range.GetLength(dimension); i++)
            {
                if (dimension == 0)
                {
                    output.Add((T) range[i, 0]);
                }
                else
                {
                    output.Add((T) range[0, i]);
                }
            }
        }
        
        return output;
    }

    /// <summary>
    /// Converts a list to object[,] which is easier to output to Excel.
    /// </summary>
    /// <param name="list">The list to convert.</param>
    /// <param name="dimension">The dimension along which to convert it i.e., 0 => column-wise, 1 => row-wise.</param>
    /// <typeparam name="T">The type.</typeparam>
    /// <returns>A 2D array for Excel.</returns>
    /// <exception cref="InvalidOperationException">Thrown if an element in the list is null.</exception>
    public static object[,] ConvertListToExcelRange<T>(List<T> list, int dimension)
    {
        if (dimension == 0)
        {
            object[,] output = new object[list.Count, 1];
            for (int i = 0; i < list.Count; i++)
            {
                output[i, 0] = list.ElementAt(i) ?? throw new InvalidOperationException();
            }
            
            return output;        
        }
        else
        {
            object[,] output = new object[1, list.Count];
            for (int i = 0; i < list.Count; i++)
            {   
                output[0, i] = list.ElementAt(i) ?? throw new InvalidOperationException();
            }
            
            return output;        
        }
    }
}
