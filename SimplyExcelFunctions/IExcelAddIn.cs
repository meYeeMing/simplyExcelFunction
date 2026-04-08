using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace SimplyExcelFunctions
{
    // Excel-DNA automatically finds any class that implements IExcelAddIn
    public class MyAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // This turns on the "Live" tooltip server when Excel starts
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            // This shuts it down cleanly when Excel closes
            IntelliSenseServer.Uninstall();
        }
    }
}