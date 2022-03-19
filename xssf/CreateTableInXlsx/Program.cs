using NPOI.OOXML.XSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace CreateTableInXlsx
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var workbook = new XSSFWorkbook();

            var sheet= (XSSFSheet)workbook.CreateSheet("Sheet1");
            //create the table in sheet
            var table = sheet.CreateTable();
            table.Name = "Test";
            var ctTable=table.GetCTTable();
            ctTable.id = 1;
            table.IsHasTotalsRow = false;
            table.DisplayName = "Table1";
            table.SetCellReferences(new NPOI.SS.Util.AreaReference("A1:C5", NPOI.SS.SpreadsheetVersion.EXCEL2007));

            table.CreateColumn(null, 0);
            table.CreateColumn(null, 1);
            table.CreateColumn(null, 2);
            table.StyleName = XSSFBuiltinTableStyleEnum.TableStyleMedium27.ToString();


            table.Style.IsShowColumnStripes = false;
            table.Style.IsShowRowStripes = true;

            //fill in the data
            for (int r = 0; r < 5; r++)
            {
                var row = sheet.CreateRow(r);
                for (int c = 0; c < 3; c++)
                {
                    var cell = row.CreateCell(c);
                    if (r == 0)
                    { //first row is for column headers
                        cell.SetCellValue("Column" + (c + 1)); //content **must** be here for table column names
                    }
                    else
                    {
                        cell.SetCellValue($"R{r + 1}C{c + 1}");
                    }
                }
            }

            using (FileStream sw = File.Create("test.xlsx"))
            {
                workbook.Write(sw);
            }
        }
    }
}
