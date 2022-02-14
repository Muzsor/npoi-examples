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
            var cttable = table.GetCTTable();
            cttable.displayName= "Table1";
            cttable.id= 1;
            cttable.@ref ="A1:C5";
            cttable.totalsRowShown =false;

            var styleInfo = cttable.tableStyleInfo = new CT_TableStyleInfo();
            styleInfo.name ="TableStyleMedium2";
            styleInfo.showColumnStripes =false;
            styleInfo.showRowStripes =true;
            cttable.tableColumns = new CT_TableColumns();
            cttable.tableColumns.tableColumn = new List<CT_TableColumn>();
            cttable.tableColumns.tableColumn.Add(new CT_TableColumn() { id = 1 });
            cttable.tableColumns.tableColumn.Add(new CT_TableColumn() { id = 2 });
            cttable.tableColumns.tableColumn.Add(new CT_TableColumn() { id = 3 });

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
