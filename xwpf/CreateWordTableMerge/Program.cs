using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.Util;
using NPOI.XWPF.UserModel;
using System;
using System.IO;

namespace CreateWordTableMerge
{
    /// <summary>
    /// https://stackoverflow.com/questions/49219092/how-to-horizontally-merge-xwpftable-using-poi-in-java
    /// </summary>
    internal class Program
    {
        static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow)
        {
            for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++)
            {
                XWPFTableCell cell = table.GetRow(rowIndex).GetCell(col);
                CT_VMerge vmerge = new CT_VMerge();
                if (rowIndex == fromRow)
                {
                    // The first merged cell is set with RESTART merge value
                    vmerge.val=ST_Merge.restart;
                }
                else
                {
                    // Cells which join (merge) the first one, are set with CONTINUE
                    vmerge.val=ST_Merge.@continue;
                    // and the content should be removed
                    for (int i = cell.Paragraphs.Count; i > 0; i--)
                    {
                        cell.RemoveParagraph(0);
                    }
                    cell.AddParagraph();
                }
                // Try getting the TcPr. Not simply setting an new one every time.
                CT_TcPr tcPr = cell.GetCTTc().tcPr;
                if (tcPr == null) tcPr = cell.GetCTTc().AddNewTcPr();
                tcPr.vMerge= vmerge;
            }
        }

        //merging horizontally by setting grid span instead of using CTHMerge
        static void mergeCellHorizontally(XWPFTable table, int row, int fromCol, int toCol)
        {
            XWPFTableCell cell = table.GetRow(row).GetCell(fromCol);
            // Try getting the TcPr. Not simply setting an new one every time.
            CT_TcPr tcPr = cell.GetCTTc().tcPr;
            if (tcPr == null) tcPr = cell.GetCTTc().AddNewTcPr();
            // The first merged cell has grid span property set
            if (tcPr.gridSpan!=null)
            {
                tcPr.gridSpan.val = (toCol - fromCol + 1).ToString();
            }
            else
            {
                tcPr.gridSpan = new CT_DecimalNumber();
                tcPr.gridSpan.val=(toCol - fromCol + 1).ToString();
            }
            // Cells which join (merge) the first one, must be removed
            for (int colIndex = toCol; colIndex > fromCol; colIndex--)
            {
                table.GetRow(row).RemoveCell(colIndex); // use only this for apache poi versions greater than 3
                                                        //table.getRow(row).getCtRow().removeTc(colIndex); // use this for apache poi versions up to 3
                                                        //table.getRow(row).removeCell(colIndex);
            }
        }
        static void Main(string[] args)
        {
            XWPFDocument document = new XWPFDocument();

            XWPFParagraph paragraph = document.CreateParagraph();
            XWPFRun run = paragraph.CreateRun();
            run.SetText("The table:");

            //create table
            XWPFTable table = document.CreateTable(3, 5);

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    table.GetRow(row).GetCell(col).SetText("row " + row + ", col " + col);
                }
            }

            //create CTTblGrid for this table with widths of the 5 columns. 
            //necessary for Libreoffice/Openoffice to accept the column widths.
            //values are in unit twentieths of a point (1/1440 of an inch)
            //first column = 1 inches width
            table.GetCTTbl().AddNewTblGrid().AddNewGridCol().w= 1 * 1440;
            //other columns (2 in this case) also each 1 inches width
            for (int col = 1; col < 5; col++)
            {
                table.GetCTTbl().tblGrid.AddNewGridCol().w=1 * 1440;
            }

            //create and set column widths for all columns in all rows
            //most examples don't set the type of the CTTblWidth but this
            //is necessary for working in all office versions
            for (int col = 0; col < 5; col++)
            {
                CT_TblWidth tblWidth =  new CT_TblWidth();
                tblWidth.w = (1 * 1440).ToString();
                tblWidth.type = ST_TblWidth.dxa;
                for (int row = 0; row < 3; row++)
                {
                    CT_TcPr tcPr = table.GetRow(row).GetCell(col).GetCTTc().tcPr;
                    if (tcPr != null)
                    {
                        tcPr.tcW=tblWidth;
                    }
                    else
                    {
                        tcPr = new CT_TcPr();
                        tcPr.tcW=tblWidth;
                        table.GetRow(row).GetCell(col).GetCTTc().tcPr= tcPr;
                    }
                }
            }

            //using the merge methods
            mergeCellVertically(table, 0, 0, 1);
            mergeCellHorizontally(table, 1, 2, 3);
            mergeCellHorizontally(table, 2, 1, 4);

            paragraph = document.CreateParagraph();
            using (FileStream fs = File.Create("create_table.docx"))
            {
                document.Write(fs);
                fs.Close();
            }

        }
    }
}
