using System;
using NPOI.SS.UserModel;

namespace ReadAndPrintData
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var workbook = WorkbookFactory.Create("data.xlsx");
            ISheet sheet = workbook.GetSheetAt(0);
            foreach (IRow row in sheet)
            {
                for (var i = 0; i < row.LastCellNum; i++)
                {
                    var cell = row.GetCell(i);
                    if (cell != null)
                    {
                        Console.Write(cell.ToString());
                        Console.Write("\t");
                    }
                }
                Console.WriteLine();
            }
            Console.Read();
        }
    }
}
