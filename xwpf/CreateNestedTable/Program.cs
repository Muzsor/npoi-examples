using NPOI.XWPF.UserModel;
using System;
using System.IO;

namespace CreateNestedTable
{
    /// <summary>
    /// https://stackoverflow.com/questions/32139125/table-inside-a-tablecell-nested-tables-with-apache-poi
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument document = new XWPFDocument();
            XWPFTable tableOne = document.CreateTable();
            XWPFTableRow tableOneRow1 = tableOne.GetRow(0);
            XWPFTableRow tableOneRow2 = tableOne.CreateRow();
            tableOneRow1.GetCell(0).SetText("Test11");
            tableOneRow1.AddNewTableCell();
            tableOneRow1.GetCell(1).SetText("Test12");
            tableOneRow2.GetCell(0).SetText("Test21");
            tableOneRow2.AddNewTableCell();

            XWPFTableCell cell = tableOneRow2.GetCell(1);
            var ctTbl = cell.GetCTTc().AddNewTbl();
            //to remove the line from the cell, you can call cell.removeParagraph(0) instead  
            cell.SetText("line1");
            cell.GetCTTc().AddNewP();

            XWPFTable tableTwo = new XWPFTable(ctTbl, cell);
            XWPFTableRow tableTwoRow1 = tableTwo.GetRow(0);
            tableTwoRow1.GetCell(0).SetText("nestedTable11");
            tableTwoRow1.AddNewTableCell();
            tableTwoRow1.GetCell(1).SetText("nestedTable12");

            using (FileStream fs = new FileStream("nestedTable.docx", FileMode.Create))
            {
                document.Write(fs);
            }
        }
    }
}
