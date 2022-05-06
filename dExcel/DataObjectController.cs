namespace dExcel;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

public static class DataObjectController
{
    private static Dictionary<string, object> dataObjects = new();

    public static string Add(string handle, object dataObject)
    {
        char[] bannedCharacters = new char[] { '@', ':' };
        if (bannedCharacters.Any(handle.Contains))
        {
            return $"Handle may not contain: {string.Join(", ", bannedCharacters)}";
        }
        if (!dataObjects.ContainsKey(handle))
        {
            dataObjects.Add(handle, dataObject);
        }
        else
        {
            dataObjects[handle] = dataObject;  
        }
        return $"@@{handle}:{DateTime.Now:HH:mm:ss}";
    }

    public static string CleanHandle(string dirtyTag)
    {
        return Regex.Match(dirtyTag, @"(?<=@@)[^:]+").Value;
    }

    public static object GetDataObject(string handle)
    {
        return dataObjects[CleanHandle(handle)];
    }
}
