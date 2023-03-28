using ExcelDna.Integration;
using System.Diagnostics;

namespace Child2
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            Debug.Print("Main AutoOpen");
        }

        public void AutoClose()
        {
        }
    }
}