namespace dExcel;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

public class AddInController : IExcelAddIn
{
    Excel.Application xlapp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
    
    public void AutoClose()
    {
        
    }

    public void AutoOpen()
    {

    }
}
