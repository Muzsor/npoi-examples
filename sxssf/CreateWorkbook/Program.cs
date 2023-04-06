using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Streaming;
using System;
using System.IO;

namespace CreateWorkbook
{
    internal class Program
    {
        static void Main(string[] args)
        {
            SXSSFWorkbook wb = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
            ISheet sh = wb.CreateSheet();
            for (int rownum = 0; rownum < 1000; rownum++)
            {
                IRow row = sh.CreateRow(rownum);
                for (int cellnum = 0; cellnum < 10; cellnum++)
                {
                    ICell cell = row.CreateCell(cellnum);
                    String address = new CellReference(cell).FormatAsString();
                    cell.SetCellValue(address);
                }
            }
            using (var fs = File.Create("test.xlsx"))
            {
                wb.Write(fs, false);
            }
        }
    }
}
