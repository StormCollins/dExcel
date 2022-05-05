namespace dExcel;

using System;
using System.Reflection;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

public static class VersionUtils
{
    public static bool TrySetWorkbookdExcelVersion(string version)
    {
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        var activeWorkbook = xlApp.ActiveWorkbook;
        var properties = (DocumentProperties)activeWorkbook.CustomDocumentProperties;
        try
        {
            foreach (DocumentProperty prop in properties)
            {
                if (string.Compare(prop.Name, "dExcel Version", StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    prop.Value += $",{version}";
                    return true;
                }
            }

            properties.Add("dExcel Version", false, MsoDocProperties.msoPropertyTypeString, version);
            return true;
        }
        catch (Exception e)
        {
            return false;
        }
    }

    public static bool TryGetWorkbookdExcelVersion(out string? version)
    {
        DocumentProperties properties;
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        var activeWorkbook = xlApp.ActiveWorkbook;
        properties = (DocumentProperties)activeWorkbook.CustomDocumentProperties;

        foreach (DocumentProperty prop in properties)
        {
            if (string.Compare(prop.Name, "dExcel Version", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                version = prop.Value.ToString();
                return true;
            }
        }
        
        version = null;
        return false;
    }
}

