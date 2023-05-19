using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace dExcel.ExcelUtils;

using Utilities;

public static class VersionUtils
{
    public static bool TrySetWorkbookdExcelVersion(string version)
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Workbook? activeWorkbook = xlApp.ActiveWorkbook;
        DocumentProperties? properties = (DocumentProperties)activeWorkbook.CustomDocumentProperties;
        try
        {
            foreach (DocumentProperty prop in properties)
            {
                if (prop.Name.IgnoreCaseEquals("dExcel Version"))
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
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Workbook? activeWorkbook = xlApp.ActiveWorkbook;
        properties = (DocumentProperties)activeWorkbook.CustomDocumentProperties;

        foreach (DocumentProperty prop in properties)
        {
            if (prop.Name.IgnoreCaseEquals("dExcel Version"))
            {
                version = prop.Value.ToString();
                return true;
            }
        }

        version = null;
        return false;
    }
}
