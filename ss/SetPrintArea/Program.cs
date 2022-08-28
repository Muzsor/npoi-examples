using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace SetPrintArea
{
    class Program
    {
        static IWorkbook workbook;
        static void Main(string[] args)
        {
            InitializeWorkbook(args);
            ISheet sheet = workbook.CreateSheet("Timesheet");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Test");
            workbook.SetPrintArea(0, "$A$1:$C$5");

            //workbook.SetPrintArea(0, "$A$1:$C$5,$E$9:$I$16");  not working in xls
            WriteToFile();
        }
        static void WriteToFile()
        {
            string filename = "timesheet.xls";
            if (workbook is XSSFWorkbook) filename += "x";
            //Write the stream data of workbook to the root directory
            using (FileStream file = new FileStream(filename, FileMode.Create))
            {
                workbook.Write(file);
            }
        }

        static void InitializeWorkbook(string[] args)
        {
            if (args.Length > 0 && args[0].Equals("-xls"))
                workbook = new HSSFWorkbook();
            else
                workbook = new XSSFWorkbook();
        }

    }
}
