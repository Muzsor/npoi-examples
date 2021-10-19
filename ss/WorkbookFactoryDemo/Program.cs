/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI Examples: https://github.com/nissl-lab/npoi-examples
 * ==============================================================*/


using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkbookFactoryDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var stream= File.OpenRead("test.xlsx"))
            {
                IWorkbook workbook = WorkbookFactory.Create(stream);
                //workbook.GetSheetAt(0).CopySheet("Folha1");
                /*for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    ISheet sheet = workbook.GetSheetAt(i);
                    Console.WriteLine(sheet.SheetName);
                }*/
                using (FileStream fileWriter = new FileStream("TestOutput.xlsx", FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                {
                    workbook.Write(fileWriter);
                }

                Console.ReadLine();
            }
        }
    }
}
