namespace dExcel;

using ExcelDna.Integration;

public static class CommonUtils
{
    public static string InFunctionWizard() => ExcelDnaUtil.IsInFunctionWizard() ? "In function wizard." : "";
}
