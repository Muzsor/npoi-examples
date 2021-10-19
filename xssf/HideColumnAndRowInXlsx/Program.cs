/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI Examples: https://github.com/nissl-lab/npoi-examples
 * ==============================================================*/

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.HideColumnAndRowInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet s = workbook.CreateSheet("Sheet1");
            for (int i = 0; i < 5; i++)
            {
                s.CreateRow(i).CreateCell(0).SetCellValue("Row "+i);
            }

            var r2 = s.GetRow(1);
            //hide Row 2
            r2.ZeroHeight = true;

            //hide column C
            s.SetColumnHidden(2, true);
            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
