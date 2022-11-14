namespace dExcel;

using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;
using FuzzySharp;
using System.Windows.Threading;
using ExcelUtils;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    // TODO: Remove blanks.

    public static IRibbonUI RibbonUi;

    public void LoadRibbon(IRibbonUI sender)
    {
        RibbonUi = sender;
    }

    public object GetImage(IRibbonControl control)
    {
        var assembly = Assembly.GetExecutingAssembly();
        return new Bitmap(
            assembly.GetManifestResourceStream($"dExcel.Resources.Icons.{control.Tag}") ??
            throw new ArgumentNullException($"Icon {control.Tag} not found in resources."));
    }

    public void OpenDashboard(IRibbonControl control)
    {
        string? dashBoardAction = null;
        var thread = new Thread(() =>
        {
            var dashboard = Dashboard.Instance;
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
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
#if DEBUG
            xlApp.Workbooks.Open(@"C:\GitLab\dExcelTools\dExcel\dExcel\Resources\Workbooks\dexcel-testing.xlsm");
#else
            xlApp.Workbooks.Open(@"C:\GitLab\dExcelTools\Releases\Current\Resources\Workbooks\dexcel-testing.xlsm");
#endif
        }
    }

    public void OpenFunctionSearch(IRibbonControl control)
    {
        string? functionName = null;
        var thread = new Thread(() =>
        {
            var functionSearch = new FunctionSearch();
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
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            ((Excel.Range)xlApp.Selection).Formula = $"=d.{functionName}()";
            ((Excel.Range)xlApp.Selection).FunctionWizard();
        }
    }

    public void InsertFunction(IRibbonControl control)
    {
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        ((Excel.Range)xlApp.Selection).Formula = $"=d.{control.Id}()";
        ((Excel.Range)xlApp.Selection).FunctionWizard();
    }

    public static IEnumerable<(string name, string description, string category)> GetCategoryMethods(string categoryName)
    {
        foreach (var method in GetExposedMethods())
        {
            if (method.category.ToUpper().Contains(categoryName.ToUpper()))
            {
                yield return method;
            }
        }
    }

    public static IEnumerable<(string name, string description, string category)> GetExposedMethods()
    {
        var methods =
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
                    => (Name: methodInfos.ElementAt(i).Name,
                        Description: excelFunctionAttribute.Description,
                        Category: excelFunctionAttribute.Category));
    }

    public string GetFunctionContent(IRibbonControl control)
    {
        var methods = GetCategoryMethods(control.Id.Replace("_", " "));
        var content = "";
        content += $"<menu xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">";
        foreach (var (name, _, _) in methods)
        {
            content +=
                $"<button " +
                $"id=\"{name.Replace(".", "")}\" " +
                $"label=\"d.{name}\" " +
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
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        xlApp.ActiveWorkbook.ApplyTheme(
            Path.Combine(
                Path.GetDirectoryName(DebugUtils.GetXllPath()),
                @"resources\workbooks\Deloitte_Brand_Theme.thmx"));
#else
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        xlApp.ActiveWorkbook.ApplyTheme(
                @"C:\GitLab\dExcelTools\Versions\CurrentDeloitte_Brand_Theme.thmx");
#endif
    }

    public void FormatTable(IRibbonControl control)
    {
        FormattingSettings? formatSettings = null;
        var thread = new Thread(() =>
        {
            var tableFormatter = TableFormatter.Instance;
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
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            bool hasTwoRowHeaders = formatSettings.Value.rowHeaderCount == 2;
            bool hasTwoColumnHeaders = formatSettings.Value.columnHeaderCount == 2;
            RangeFormatUtils.SetColumnAndRowHeaderBasedTableFormatting(hasTwoRowHeaders, hasTwoColumnHeaders);
        }
        else if (formatSettings is { columnHeaderCount: > 0 })
        {
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            bool hasTwoHeaders = formatSettings.Value.columnHeaderCount == 2;
            RangeFormatUtils.SetColumnHeaderBasedTableFormatting(hasTwoHeaders);
        }
        else if (formatSettings is { rowHeaderCount: > 0 })
        {
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            bool hasTwoHeaders = formatSettings.Value.rowHeaderCount == 2;
            RangeFormatUtils.SetRowHeaderBasedTableFormatting(hasTwoHeaders);
        }
    }
    

    /// <summary>
    /// Calculates the selected Excel range.
    /// </summary>
    /// <param name="control">Ribbon control.</param>
    public void CalculateRange(IRibbonControl control)
    {
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        ((Excel.Range)xlApp.Selection).Calculate();
    }

    public void ApplyLogicFormatting(IRibbonControl control)
    {
        RangeFormatUtils.SetRangeConditionalLogicFormatting();
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
                wb.VBProject.References.Remove(reference);
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
}
