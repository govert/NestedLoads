using ExcelDna.Integration;
using System.Diagnostics;

namespace Child2
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            Debug.Print("Child2 AutoOpen");
        }

        public void AutoClose()
        {
        }
    }
}