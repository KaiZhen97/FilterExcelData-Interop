using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
//using Workbook = Microsoft.Office.Interop.Excel.Workbook;
//using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
//using Range = Microsoft.Office.Interop.Excel.Range;
//using Delete = Microsoft.Office.Interop.Excel.XlDeleteShiftDirection;


namespace filterData
{
    class Program
    {
        static void Main(string[] args)
        {
            var excel = new Application();
            Workbook workbook = excel.Workbooks.Open(@"C:\Users\kaizhen.goh\source\report\CrystalReportViewer1 (6).xlsx");
            Worksheet worksheet = workbook.Sheets[1];
            //Range range = worksheet.UsedRange; // include column name
            Range range = worksheet.Range["A2:P27511"];

            range.AutoFilter(4, "<>YN3210X");
            range.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
            //range.Delete(XlDeleteShiftDirection.xlShiftUp);

            workbook.SaveAs(@"C:\Users\kaizhen.goh\source\report\NEW\text.xlsx");
            excel.Visible = true;
        }
    }
}
