namespace dExcel;

using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;
using FuzzySharp;
using System.Windows.Threading;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    // TODO: Switch off page breaks easily.
    // TODO: Link to Wiki for functions.
    // TODO: Link to GitLab.
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
            assembly.GetManifestResourceStream($"dExcel.resources.icons.{control.Tag}") ??
            throw new ArgumentNullException($"Icon {control.Tag} not found in resources."));
    }

    public void OpenDashboard(IRibbonControl control)
    {
        var thread = new Thread(() =>
        {
            var dashboard = Dashboard.Instance;
            dashboard.Show();
            dashboard.Closed += (sender2, e2) => dashboard.Dispatcher.InvokeShutdown();
            Dispatcher.Run();
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
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

    public void ClearFormatting(IRibbonControl control) => CellFormatUtils.ClearFormatting();

    public void FormatTable(IRibbonControl control) => CellFormatUtils.FormatTable();

    /// <summary>
    /// Calculates the selected Excel range.
    /// </summary>
    /// <param name="control">Ribbon control.</param>
    public void CalculateRange(IRibbonControl control)
    {
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        ((Excel.Range)xlApp.Selection).Calculate();
    }

    // TODO: Move these to a separate class.
    private StringBuilder allTemplates = new("Hello,Hi,XiXi,Hoho,XiXe,Hiddy");
    private StringBuilder templates = new("Hello,Hi,XiXi,Hoho,XiXe,Hiddy");
    private static int templateCount = 6;

    public int GetTemplateSearchCount(IRibbonControl control)
    {
        return templateCount;
    }

       
    public void WrapUpAudit(IRibbonControl control)
    {
        // Remove review notes
        // Get rid of headings.
        // Hide formuale
        // Hide cell names
        // Lock sheets and protect with password
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
            if (worksheet.ProtectContents == true)
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
}
