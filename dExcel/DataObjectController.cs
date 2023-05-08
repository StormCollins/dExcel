using System.Text.RegularExpressions;
using dExcel.Utilities;

namespace dExcel;

public sealed class DataObjectController
{
    private readonly Dictionary<string, object> _dataObjects;

    private DataObjectController()
    {
        _dataObjects = new Dictionary<string, object>();
    }

    private static DataObjectController _instance = new();

    public static DataObjectController Instance
    {
        get
        {
            if (_instance == null)
            {
                _instance = new DataObjectController();
            }

            return _instance;
        }
    }

    public string Add(string handle, object dataObject)
    {
        char[] bannedCharacters = { '@', ':', ',', ';', '\\', '/' };
        if (bannedCharacters.Any(handle.Contains))
        {
            return CommonUtils.DExcelErrorMessage($"Handle may not contain following: {string.Join(", ", bannedCharacters)}");
        }
        
        if (!_dataObjects.ContainsKey(handle))
        {
            _dataObjects.Add(handle, dataObject);
        }
        else
        {
            _dataObjects[handle] = dataObject;  
        }
        return $"@@{handle}::{DateTime.Now:HH:mm:ss}";
    }

    private static string CleanHandle(string dirtyTag)
    {
        return Regex.Match(dirtyTag, @"(?<=@@)[^:]+").Value;
    }

    public object GetDataObject(string handle)
    {
        return _dataObjects[CleanHandle(handle)];
    }
}
