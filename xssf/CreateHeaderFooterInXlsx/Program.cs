/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI Examples: https://github.com/nissl-lab/npoi-examples
 * ==============================================================*/

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace CreateHeaderFooterInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            using (IWorkbook workbook = new XSSFWorkbook())
            {
                ISheet s1 = workbook.CreateSheet("Sheet1");
                s1.CreateRow(0).CreateCell(1).SetCellValue(123);

                //set header text
                s1.Header.Left = HSSFHeader.Page;   //Page is a static property of HSSFHeader and HSSFFooter
                s1.Header.Center = "This is a test sheet";
                //set footer text
                s1.Footer.Left = "Copyright Nissl Lab";
                s1.Footer.Right = "created by NPOI team";
                using (FileStream sw = File.Create("test.xlsx"))
                {
                    workbook.Write(sw, false);
                }
            }
        }
    }
}
