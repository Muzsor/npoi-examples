using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace NPOI.Examples.XSSF.CreateEmptyWorkbook
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            workbook.CreateSheet("Sheet 1");
            workbook.CreateSheet("Sheet 2");
            workbook.CreateSheet("Sheet 3");
                        
            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();

            Console.WriteLine("File 'test.xls' generated");
            Console.ReadLine();
        }
    }
}
