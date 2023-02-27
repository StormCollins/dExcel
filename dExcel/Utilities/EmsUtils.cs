namespace dExcel.Utilities;

using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

public static class EmsUtils
{
    [ExcelFunction(
        Name = "d.EmsUtils_FixEmsLinks",
        Description = "Fixes links to workbooks that users have linked to in EMS.",
        Category = "∂Excel: EMS Utils")]
    public static void FixEmsLinks()
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        // The workbook who's links need to be fixed.
        Excel.Workbook targetWorkbook = xlApp.ActiveWorkbook;
        string targetWorkbookName = targetWorkbook.Name;
        List<string> sourceWorkbookNames = new List<string>();
        int openWorkbooksCount = xlApp.Workbooks.Count;
        List<string> openWorkbookNames = new();

        for (int i = 1; i <= openWorkbooksCount; i++)
        {
            if (targetWorkbookName != xlApp.Workbooks[i].Name)
            {
                sourceWorkbookNames.Add(xlApp.Workbooks[i].Name);
            }

            openWorkbookNames.Add(xlApp.Workbooks[i].Name);
        }

        foreach (Excel.Worksheet currentTargetWorksheet in targetWorkbook.Worksheets)
        {
            Excel.Range xlUsedRange = currentTargetWorksheet.UsedRange;

            foreach (string sourceWorkbookName in sourceWorkbookNames)
            {
                if (currentTargetWorksheet.Name == "Fix EMS Links Utils")
                {
                    var x = 2;
                }

                Dictionary<string?, Excel.Range> allRanges = FindAllRangesWithText(sourceWorkbookName);

                foreach (KeyValuePair<string?, Excel.Range> range in allRanges)
                {
                    string oldFormula = currentTargetWorksheet.Range[range.Key].Formula;
                    string newFormula = Regex.Replace(oldFormula, @"(?<=').+\\", "");
                    currentTargetWorksheet.Range[range.Key].FormulaArray = newFormula;
                }
            }
            
            Dictionary<string?, Excel.Range> FindAllRangesWithText(string stringToFind)
            {
                Dictionary<string?, Excel.Range> allRanges = new();
                Excel.Range? foundRange =
                    xlUsedRange.Find(
                            What: stringToFind,
                            LookIn: Excel.XlFindLookIn.xlFormulas,
                            LookAt: Excel.XlLookAt.xlPart,
                            SearchOrder: Excel.XlSearchOrder.xlByRows,
                            SearchDirection: Excel.XlSearchDirection.xlNext,
                            MatchCase: false,
                            SearchFormat: false);    
                
                do
                {
                    if (foundRange != null)
                    {
                        allRanges[foundRange.Address] = foundRange;
                        foundRange = xlUsedRange.FindNext(foundRange);
                    }
                    else
                    {
                        break;
                    }
                } while (!allRanges.ContainsKey(foundRange?.Address)); 
                
                return allRanges;
            }
        }
    }
}
