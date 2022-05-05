namespace dExcel;

using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;
using FuzzySharp;

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

    public void GetDiagnostics(IRibbonControl control)
    {
        var diagnostics = Diagnostics.Instance;
        diagnostics.Show();
    }

    public void OpenFunctionSearch(IRibbonControl control)
    {
        var functionSearch = new FunctionSearch();
        functionSearch.Show();
    }

    public void InsertFunction(IRibbonControl control)
    {
        var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        ((Excel.Range)xlApp.Selection).Formula = $"=d.{control.Id}()";
        ((Excel.Range)xlApp.Selection).FunctionWizard();
    }

    public List<(string name, string description, string category)> GetCategoryMethods(string categoryName)
    {
        var methods =
            Assembly
                .GetExecutingAssembly()
                .GetTypes()
                .SelectMany(x => x.GetMethods())
                .Where(y => y.GetCustomAttributes(typeof(ExcelFunctionAttribute), false).Length > 0)
                .Where(z =>
                    z.GetCustomAttribute(typeof(ExcelFunctionAttribute)) is not ExcelFunctionAttribute
                    || (z.GetCustomAttribute(typeof(ExcelFunctionAttribute)) as ExcelFunctionAttribute)
                    .Category?.ToUpper().Contains(categoryName.ToUpper()) == true);

        var methodInfos = methods as MethodInfo[] ?? methods.ToArray();
        return methodInfos.Select((t, i)
            => (ExcelFunctionAttribute)methodInfos
                .ElementAt(i)
                .GetCustomAttribute(typeof(ExcelFunctionAttribute)))
                .Select((excelFunctionAttribute, i)
                    => (methodInfos.ElementAt(i).Name,
                        excelFunctionAttribute.Description,
                        excelFunctionAttribute.Category))
                .ToList();
    }

    public List<(string name, string description, string category)> GetExposedMethods()
    {
        var methods =
            Assembly
                .GetExecutingAssembly()
                .GetTypes()
                .SelectMany(x => x.GetMethods())
                .Where(y => y.GetCustomAttributes(typeof(ExcelFunctionAttribute), false).Length > 0)
                .Where(z => z.GetCustomAttribute(typeof(ExcelFunctionAttribute)) is not ExcelFunctionAttribute);

        var methodInfos = methods as MethodInfo[] ?? methods.ToArray();
        return methodInfos.Select((t, i)
            => (ExcelFunctionAttribute)methodInfos
                .ElementAt(i)
                .GetCustomAttribute(typeof(ExcelFunctionAttribute)))
                .Select((excelFunctionAttribute, i)
                    => (methodInfos.ElementAt(i).Name,
                        excelFunctionAttribute.Description,
                        excelFunctionAttribute.Category))
                .ToList();
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


    // TODO: Move these to a separate class.
    private StringBuilder allTemplates = new("Hello,Hi,XiXi,Hoho,XiXe,Hiddy");
    private StringBuilder templates = new("Hello,Hi,XiXi,Hoho,XiXe,Hiddy");
    private static int templateCount = 6;

    public int GetTemplateSearchCount(IRibbonControl control)
    {
        return templateCount;
    }

    public string GetTemplateSearchItemLabel(IRibbonControl control, int index)
    {
        string[] separators = new string[1];
        separators[0] = ",";
        String[] newstring = templates.ToString().Split(separators, StringSplitOptions.None);
        String? str = newstring[index];
        return str ?? "";
    }
    
    public void EditBoxTextChanged(IRibbonControl control, string text)
    {
        if (text == "")
        {
            templates = allTemplates;
            templateCount = 6;
        }
        else
        {
            var matches = Process.ExtractTop(text, allTemplates.ToString().Split(','));

            templates = new("");
            templateCount = 0;

            foreach (var match in matches)
            {
                templates.Append($",{match.Value}");
                templateCount++;
            }
            templates = new(templates.ToString().TrimStart(','));
            
        }
        RibbonUi.InvalidateControl("TemplateSearch");
    }

    public void Add(IRibbonControl control)
    {
        templateCount++;
        templates.Append("||Item" + templateCount.ToString());
        RibbonUi.InvalidateControl("TemplateSearch");
    }
}
