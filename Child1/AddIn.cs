using ExcelDna.Integration;
using System.Diagnostics;
using System.IO;

namespace Child1
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            Debug.Print("Child1 AutoOpen - Enter");

            var xlApp = ExcelDnaUtil.Application as Microsoft.Office.Interop.Excel.Application;
            var addInDir = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            xlApp.RegisterXLL(Path.Combine(addInDir, "Child2-AddIn64.xll"));

            Debug.Print("Child1 AutoOpen - Exit");
        }

        public void AutoClose()
        {
        }
    }
}
