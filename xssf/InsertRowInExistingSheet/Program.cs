using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace InsertRowInExistingSheet
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook;
            using (var fs = new FileStream("test.xlsx", FileMode.Open))
            {
                workbook = new XSSFWorkbook(fs);
            }
            var sheet = workbook.GetSheetAt(0);
            sheet.ShiftRows(4, 5, 2);
            sheet.ShiftRows(14, 16, 5);
            using (FileStream sw = File.Create("output1.xlsx"))
            {
                workbook.Write(sw);
            }
        }
    }
}
