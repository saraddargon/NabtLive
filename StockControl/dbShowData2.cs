using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telerik.WinControls.UI;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
namespace StockControl
{
    public static class dbShowData2
    {
        public static void InsertToExcel(ref Excel.Worksheet exc, string Column, dynamic Values)
        {
            try
            {
                Excel.Range refs = exc.get_Range(Column);
                refs.Value2 = Values;
                if (Values.Contains("Æ"))
                {
                    int addint = Values.IndexOf("Æ");
                    refs.Characters[addint, 2].Font.Name = "Symbol";//                          
                }
            }
            catch { }
        }
    }
}
