namespace dExcel;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

public static class DataObjectController
{
    private static readonly Dictionary<string, object> DataObjects = new();

    public static string Add(string handle, object dataObject)
    {
        char[] bannedCharacters = new char[] { '@', ':' };
        if (bannedCharacters.Any(handle.Contains))
        {
            return $"{CommonUtils.DExcelErrorPrefix} Handle may not contain following: {string.Join(", ", bannedCharacters)}";
        }
        if (!DataObjects.ContainsKey(handle))
        {
            DataObjects.Add(handle, dataObject);
        }
        else
        {
            DataObjects[handle] = dataObject;  
        }
        return $"@@{handle}::{DateTime.Now:HH:mm:ss}";
    }

    public static string CleanHandle(string dirtyTag)
    {
        return Regex.Match(dirtyTag, @"(?<=@@)[^:]+").Value;
    }

    public static object GetDataObject(string handle)
    {
        return DataObjects[CleanHandle(handle)];
    }
}
