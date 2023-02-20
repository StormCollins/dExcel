namespace dExcel;

using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Text;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration;
using ExcelUtils;
using FuzzySharp;
using SkiaSharp;
using Utilities;
using WPF;

/// <summary>
/// Used to control the actions and behaviour of the ribbon of the add-in.
/// </summary>
[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    /// <summary>
    /// This maps category names, in the ExcelFunction attribute, to the relevant ribbon function drop down menu.
    ///
    /// For instance, the function 'd.Equity_BlackScholes' is in the category '∂Excel: Equities' (the '∂Excel: ' part is
    /// ignored in the mapping). So obviously it is added to the drop down menu for 'Equities'.
    ///
    /// This also allows us to map multiple function categories to the same drop down menu e.g., anything involving
    /// cross currency swaps, 'XCCY', can be mapped to 'InterestRates' and 'FX' simultaneously.
    /// </summary>
    private static Dictionary<string, List<string>> _excelFunctionCategoriesToRibbonLabels = new()
    {
        ["DATE"] = new List<string> { "Date" },
        ["MATH"] = new List<string> { "Math" },
        ["STATS"] = new List<string> { "Stats" },
        ["CREDIT"] = new List<string> { "Credit" },
        ["COMMODITIES"] = new List<string> { "Commodities" },
        ["EQUITIES"] = new List<string> { "Equities" },
        ["FX"] = new List<string> { "FX", "XCCY"},
        ["INTERESTRATES"] = new List<string> { "CurveUtils", "HullWhite", "Interest Rates", "XCCY"},
        ["OTHER"] = new List<string> { "Debug", "Test"},
    };

    private static IRibbonUI? _ribbonUi;
    
    public void LoadRibbon(IRibbonUI? sender)
    {
        _ribbonUi = sender;
    }

    public object GetImage(IRibbonControl control)
    {
        Assembly assembly = Assembly.GetExecutingAssembly();
        return new Bitmap(
            assembly.GetManifestResourceStream($"dExcel.Resources.Icons.{control.Tag}") ??
            throw new ArgumentNullException($"Icon {control.Tag} not found in resources."));
    }

    public void OpenDashboard(IRibbonControl control)
    {
        string? dashBoardAction = null;
        Thread thread = new(() =>
        {
            Dashboard dashboard = Dashboard.Instance;
            dashboard.Show();
            dashboard.Closed += (sender2, e2) => dashboard.Dispatcher.InvokeShutdown();
            Dispatcher.Run();
            dashBoardAction = dashboard.DashBoardAction;
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (string.Compare(dashBoardAction, "OpenTestingWorkbook", true) == 0)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
#if DEBUG
            Excel.Workbook wb = xlApp.Workbooks.Open(@"C:\GitLab\dExcelTools\dExcel\dExcel\Resources\Workbooks\dexcel-testing.xlsm");
#else
            Excel.Workbook wb = xlApp.Workbooks.Open(@"C:\GitLab\dExcelTools\Releases\Current\Resources\Workbooks\dexcel-testing.xlsm", ReadOnly: true);
#endif
            Excel.Worksheet ws = wb.Worksheets["Summary"];
            ws.Activate();
            ws.Cells[1, 1].Select();
            xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;
        }
    }

    public void OpenFunctionSearch(IRibbonControl control)
    {
        string? functionName = null;
        Thread thread = new(() =>
        {
            FunctionSearch functionSearch = new();
            functionSearch.SearchTerm.Focus();
            functionSearch.Show();
            functionSearch.Closed += (sender2, e2) => functionSearch.Dispatcher.InvokeShutdown();
            Dispatcher.Run();
            functionName = functionSearch.FunctionName;
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (functionName != null)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            ((Excel.Range)xlApp.Selection).Formula = $"={functionName}()";
            ((Excel.Range)xlApp.Selection).FunctionWizard();
        }
    }

    public void InsertFunction(IRibbonControl control)
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        ((Excel.Range)xlApp.Selection).Formula = $"={control.Id}()";
        ((Excel.Range)xlApp.Selection).FunctionWizard();
    }

    public void CreateLinkToSheet(IRibbonControl control)
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        List<string> sheetNames = new();
        foreach (Excel.Worksheet sheet in xlApp.Sheets)
        {
            sheetNames.Add(sheet.Name);
        }

        string searchText = ((Excel.Range)xlApp.Selection).Value2;
        string matchedSheet = Process.ExtractTop(searchText, sheetNames).First(x => x.Score > 90).Value;
        ((Excel.Worksheet)xlApp.ActiveSheet).Hyperlinks.Add(xlApp.Selection, "", $"'{matchedSheet}'!A1");
    }

    /// <summary>
    /// Creates a hyperlink in the selected cell to a cell with the same content but which also has a heading style.
    /// For example, if the selected cell has the content "Test" and there is another cell, styled as a heading, also with
    /// the content "Test", then this function will create a link in the selected cell to the heading cell.
    ///
    /// Note this function can handle multiple cells as well.
    /// </summary>
    /// <param name="control">The ribbon control.</param>
    public void CreateHyperlinksToHeadingsInCurrentSheet(IRibbonControl control)
    {
        HyperLinkUtils.CreateHyperlinkToHeadingInSheet();
    }
    
    public void CreateLinksToHeadingsInOtherSheets(IRibbonControl control)
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Workbook wb = xlApp.ActiveWorkbook;
        foreach (Excel.Worksheet ws in wb.Worksheets)
        {
            Excel.Range usedRange = ws.UsedRange;
            Dictionary<string, List<string>> headings = new();

            for (int i = 1; i < usedRange.Rows.Count + 1; i++)
            {
                for (int j = 1; j < usedRange.Columns.Count + 1; j++)
                {
                    Excel.Range currentCell = usedRange.Cells[i, j];
                    string style = ((Excel.Style)currentCell.Style).Value.ToUpper();
                    if (style.Contains("HEADING"))
                    {
                        try
                        {
                            if (!headings.ContainsKey(currentCell.Value2))
                            {
                                headings.Add(currentCell.Value2, new List<string>() {$"'{currentCell.Name}'!{currentCell.Address}"});
                            }
                            else
                            {
                                headings[currentCell.Value2].Add($"'{currentCell.Name}'!{currentCell.Address}");
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }

            Excel.Range selectedRange = ((Excel.Range)xlApp.Selection);
            List<string> failedHyperlinks = new();

            for (int i = 1; i < selectedRange.Rows.Count + 1; i++)
            {
                string searchText = "";
                for (int j = 1; j < selectedRange.Columns.Count + 1; j++)
                {
                    Excel.Range currentCell = selectedRange.Rows[i].Columns[j];
                    if (currentCell.Value2 != null && currentCell.Value2?.ToString() != "" &&
                        currentCell.Value2?.ToString() != ExcelEmpty.Value.ToString())
                    {
                        searchText = selectedRange.Rows[i].Columns[j].Value2;
                    }
                }

                try
                {
                    string matchedHeading = Process.ExtractTop(searchText, headings.Keys).First(x => x.Score > 90).Value;
                    ((Excel.Worksheet)xlApp.ActiveSheet).Hyperlinks.Add(selectedRange.Rows[i], "",
                        headings[matchedHeading]);
                }
                catch (Exception)
                {
                    failedHyperlinks.Add(searchText);
                }
            }

            if (failedHyperlinks.Count > 0)
            {
                Alert alert = new()
                {
                    AlertCaption =
                    {
                        Text = "Warning: Section Headings Not Found"
                    },
                    AlertBody =
                    {
                        Text = "The following headings could not be found: \n  •" + string.Join("\n  • ", failedHyperlinks)
                    },
                };

                alert.Show();
            }

        }
    }
    

    public static IEnumerable<(string name, string description, string category)> 
        GetCategoryMethods(List<string> categoryNames)
    {
        foreach ((string name, string description, string category) method in GetExposedMethods())
        {
            foreach (string categoryName in categoryNames)
            {
                if (method.category.ToUpper().Contains(categoryName.ToUpper()))
                {
                    yield return method;
                }
            }
        }
    }

    public static IEnumerable<(string name, string description, string category)> GetExposedMethods()
    {
        IEnumerable<MethodInfo> methods =
            Assembly
                .GetExecutingAssembly()
                .GetTypes()
                .SelectMany(x => x.GetMethods())
                .Where(y => y.GetCustomAttributes(typeof(ExcelFunctionAttribute), false).Length > 0)
                .Where(z => z.GetCustomAttribute(typeof(ExcelFunctionAttribute)) is ExcelFunctionAttribute);

        var methodInfos = methods as MethodInfo[] ?? methods.ToArray();
        return methodInfos.Select((t, i)
                => (ExcelFunctionAttribute)methodInfos
                    .ElementAt(i)
                    .GetCustomAttribute(typeof(ExcelFunctionAttribute)))
            .Select((excelFunctionAttribute, i)
                => (Name: excelFunctionAttribute.Name,
                    Description: excelFunctionAttribute.Description,
                    Category: excelFunctionAttribute.Category));
    }

    public string GetFunctionContent(IRibbonControl control)
    {
        List<string> methodIds = _excelFunctionCategoriesToRibbonLabels[control.Id.ToUpper()];
        IEnumerable<(string name, string description, string category)> methods = GetCategoryMethods(methodIds);
        methods = methods.OrderBy(x => x.name);
        string content = "";
        content += $"<menu xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">";
        string currentSubcategory = "";
        foreach (var (name, _, _) in methods)
        {
            var previousSubcategory = currentSubcategory;
            currentSubcategory =
                Regex.Match(name, @"(?<=d\.)[^_]+", RegexOptions.Compiled | RegexOptions.IgnoreCase).Value;

            if (previousSubcategory != "" && previousSubcategory != currentSubcategory)
            {
                content += $"<menuSeparator id='separator{currentSubcategory}'/>";
            }

            content +=
                $"<button " +
                $"id=\"{name}\" " +
                $"label=\"{name}\" " +
                $"onAction=\"InsertFunction\" />";

        }

        content += "</menu>";
        return content;
    }



    public string GetTemplateContent(IRibbonControl control)
    {
        string path = @"\\ZAJNB010\FSA Valuations\FSA Valuations\Model Validation";
        var content = "";

        return path;
    }

    /// <summary>
    /// Sets the official ISO-8601 date format in the selected cell i.e., yyyy-MM-dd.
    /// This is quicker than having to navigate to it via the standard Excel menus.
    /// </summary>
    public void SetIso8601DateFormatting(IRibbonControl control)
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        ((Excel.Range)xlApp.Selection).NumberFormat = "yyyy-mm-dd;@";
    }

    /// <summary>
    /// Removes the formatting from a table (or rather contiguous region) in Excel.
    /// </summary>
    /// <param name="control">The ribbon control.</param>
    public void ClearTableFormatting(IRibbonControl control) => RangeFormatUtils.ClearRangeFormatting();

    /// <summary>
    /// Loads and switches Excel to the standard Deloitte Excel theme i.e. it loads the relevant .thmx file.
    /// </summary>
    /// <param name="control"></param>
    public void LoadDeloitteTheme(IRibbonControl control)
    {
#if DEBUG
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        xlApp.ActiveWorkbook.ApplyTheme(
            Path.Combine(
                Path.GetDirectoryName(DebugUtils.GetXllPath()) ?? string.Empty,
                @"resources\workbooks\Deloitte_Brand_Theme.thmx"));
#else
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        xlApp.ActiveWorkbook.ApplyTheme(
                @"C:\GitLab\dExcelTools\Releases\Current\Deloitte_Brand_Theme.thmx");
#endif
    }

    public void FormatTable(IRibbonControl control)
    {
        FormattingSettings? formatSettings = null;
        Thread thread = new(() =>
        {
            TableFormatter tableFormatter = TableFormatter.Instance;
            tableFormatter.Show();
            tableFormatter.Closed += (sender2, e2) => tableFormatter.Dispatcher.InvokeShutdown();
            Dispatcher.Run();
            formatSettings = tableFormatter.FormatSettings;
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (formatSettings is { rowHeaderCount: > 0, columnHeaderCount: > 0 })
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            bool hasTwoRowHeaders = formatSettings.Value.rowHeaderCount == 2;
            bool hasTwoColumnHeaders = formatSettings.Value.columnHeaderCount == 2;
            RangeFormatUtils.SetColumnAndRowHeaderBasedTableFormatting(hasTwoRowHeaders, hasTwoColumnHeaders);
        }
        else if (formatSettings is { columnHeaderCount: > 0 })
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            bool hasTwoHeaders = formatSettings.Value.columnHeaderCount == 2;
            RangeFormatUtils.SetVerticallyAlignedTableFormatting(hasTwoHeaders);
        }
        else if (formatSettings is { rowHeaderCount: > 0 })
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            bool hasTwoHeaders = formatSettings.Value.rowHeaderCount == 2;
            RangeFormatUtils.SetHorizontallyAlignedTableFormatting(hasTwoHeaders);
        }
    }
    

    /// <summary>
    /// Calculates the selected Excel range.
    /// </summary>
    /// <param name="control">Ribbon control.</param>
    public void CalculateRange(IRibbonControl control)
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        ((Excel.Range)xlApp.Selection).Calculate();
    }

    public void ApplyTestUtilsFormatting(IRibbonControl control)
    {
        RangeFormatUtils.SetConditionalTestUtilsFormatting();
    }

    public void ApplyPrimaryHeaderFormatting(IRibbonControl control)
    {
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        RangeFormatUtils.SetPrimaryTitleRangeFormatting(xlApp.Selection);
    }

    public void ApplySecondaryHeaderFormatting(IRibbonControl control)
    {
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        RangeFormatUtils.SetSecondaryTitleRangeFormatting(xlApp.Selection);
    }

    // TODO: Move these to a separate class.
    private StringBuilder allTemplates = new("Hello,Hi,XiXi,Hoho,XiXe,Hiddy");
    private StringBuilder templates = new("Hello,Hi,XiXi,Hoho,XiXe,Hiddy");
    private static int templateCount = 6;

    public int GetTemplateSearchCount(IRibbonControl control)
    {
        return templateCount;
    }

    /// <summary>
    /// Opens an (EMS) audit file that was previously "wrapped up" using the <see cref="WrapUpAudit(IRibbonControl)"/> function.
    /// </summary>
    /// <param name="control">The ribbon control.</param>
    public void OpenAuditFile(IRibbonControl control)
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

        foreach (Excel.Worksheet worksheet in xlApp.ActiveWorkbook.Worksheets)
        {
            if (worksheet.ProtectContents)
            {
                worksheet.Unprotect("asterix");
            }

            xlApp.ActiveWindow.DisplayHeadings = true;
        }

        foreach (Excel.Name name in xlApp.ActiveWorkbook.Names)
        {
            name.Visible = true;
        }
    }

    /// <summary>
    /// This wraps up an audit Excel file by:
    ///     <list type="number">
    ///         <item> Deleting all review notes.</item>
    ///         <item> Hiding the row and column headings.</item>
    ///         <item> Hiding formulae.</item>
    ///         <item> Hiding range names. </item>
    ///         <item> Locking and password protecting cells.</item>
    ///         <item> Removing Matlab ExcelLink references.</item>
    ///         <item> Password protecting VBA code.</item>
    ///     </list>
    /// </summary>
    /// <param name="control">The ribbon control.</param>
    public void WrapUpAudit(IRibbonControl control)
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
    
        // Delete review notes.
        // It is unclear where the workbook WPExcel.xls is located but EMS users seem to have access to it.
        xlApp.Application.Run("WPEXCEL.XLS!DPWPReviewNotesDeleteAll");
    
        foreach (Excel.Worksheet worksheet in xlApp.ActiveWorkbook.Worksheets)
        {
            worksheet.Activate();
            int lastRowOfUsedRange = worksheet.UsedRange.Rows.Count;
            int lastColumnOfUsedRange = worksheet.UsedRange.Columns.Count;
            xlApp.ActiveWindow.DisplayHeadings = false;
            
            if (!worksheet.ProtectContents)
            {
                worksheet.UsedRange.Cells.FormulaHidden = true;
                worksheet.UsedRange.Cells.Locked = true;
                worksheet.Protect("asterix");
            }
        }
    
        // Remove Matlab references.
        // TODO: This can be deprecated when we no longer use Matlab and ExcelLink.
        Excel.Workbook wb = xlApp.ActiveWorkbook;
        foreach (VBIDE.Reference reference in wb.VBProject.References)
        {
            if (reference.Name.Contains("ExcelLink", StringComparison.CurrentCultureIgnoreCase) ||
                reference.Name.Contains("SpreadsheetLink", StringComparison.CurrentCultureIgnoreCase))
            {
                wb.VBProject.References.Remove((Microsoft.Vbe.Interop.Reference)reference);
            }
        }

        foreach (VBIDE.VBComponent component in wb.VBProject.VBComponents)
        {
            VBIDE.CodeModule codeModule = component.CodeModule;

            for (int i = 1; i < codeModule.CountOfLines; i++)
            {
                string line = codeModule.Lines[i, 1];
                if (line.Contains("#Const oExcelLink = 1"))
                {
                    codeModule.ReplaceLine(i, line.Replace("#Const oExcelLink = 1", "#Const oExcelLink = 0"));
                }
            }
        }

    // Password protect the VBA code.
    // TODO: See if this can be simplified.
    xlApp.Application.ScreenUpdating = false;
    
     string breakIt = "%{F11}%TE+{TAB}{RIGHT}%V{+}{TAB}";
     foreach (VBIDE.Window window in wb.VBProject.VBE.Windows)
     {
         if (window.Caption.Contains('('))
         {
             window.Close();
         }
     }
    wb.Activate();
         
          xlApp.Application.OnKey("%{F11}");
          SendKeys.SendWait(breakIt + "asterix" + "{TAB}" + "asterix" + "~%{F11}");
          xlApp.Application.ScreenUpdating = true;
          wb.Activate();
    }

    public void ViewObjectChart(IRibbonControl control)
    {
        SkiaSharp.Views.WPF.WPFExtensions.ToColor(SKColor.Empty);
        CurvePlotter curvePlotter = CurvePlotter.Instance;
        curvePlotter.Show();
    }

    public void FixEMSLinks(IRibbonControl ribbonControl)
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Workbook xlActiveWorkbook = xlApp.ActiveWorkbook;

        int numberOfOpenWorkbooks = xlApp.Workbooks.Count;
        List<string> xlOpenWorkbooks = new();
        for (int i = 1; i <= numberOfOpenWorkbooks; i++)
        {
            xlOpenWorkbooks.Add(xlApp.Workbooks[i].Name);
        }

        foreach (Excel.Worksheet xlWorksheet in xlActiveWorkbook.Worksheets)
        {
            Excel.Range xlUsedRange = xlWorksheet.UsedRange;

            foreach (string xlOpenWorkbook in xlOpenWorkbooks)
            {
                if (xlOpenWorkbook != xlActiveWorkbook.Name)
                {
                    Excel.Range workbookFind = ExcelFind(xlOpenWorkbook);

                    while (workbookFind is not null)
                    {
                        string filePath = 
                            Regex.Match(xlWorksheet.Range[workbookFind.Address].Formula, @"(?<==')[^\[]+").Value;

                        xlUsedRange.Replace(
                            What: filePath,
                            Replacement: "",
                            LookAt: Excel.XlLookAt.xlPart,
                            SearchOrder: Excel.XlSearchOrder.xlByRows,
                            MatchCase: true,
                            SearchFormat: false,
                            ReplaceFormat: false);

                        workbookFind = ExcelFind(xlOpenWorkbook);
                    }
                }
            }

            Excel.Range ExcelFind(string stringToFind)
                => xlUsedRange.Find(
                    What: "\\[" + stringToFind + "]",
                    LookIn: Excel.XlFindLookIn.xlFormulas,
                    LookAt: Excel.XlLookAt.xlPart,
                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Excel.XlSearchDirection.xlNext,
                    MatchCase: true,
                    SearchFormat: false);
        }
    }
}
