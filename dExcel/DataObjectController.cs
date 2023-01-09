namespace dExcel;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Utilities;

public static class DataObjectController
{
    private static readonly Dictionary<string, object> DataObjects = new();

    public static string Add(string handle, object dataObject)
    {
        char[] bannedCharacters = { '@', ':' };
        if (bannedCharacters.Any(handle.Contains))
        {
            return CommonUtils.DExcelErrorMessage($"Handle may not contain following: {string.Join(", ", bannedCharacters)}");
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

    private static string CleanHandle(string dirtyTag)
    {
        return Regex.Match(dirtyTag, @"(?<=@@)[^:]+").Value;
    }

    public static object GetDataObject(string handle)
    {
        return DataObjects[CleanHandle(handle)];
    }
}
