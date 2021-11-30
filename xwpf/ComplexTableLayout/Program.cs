/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI Examples: https://github.com/nissl-lab/npoi-examples
 * ==============================================================*/

using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.IO;

namespace ComplexTableLayout
{
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFTable table1 = doc.CreateTable(3, 3);
            var tblLayout1 = table1.GetCTTbl().tblPr.AddNewTblLayout();
            tblLayout1.type = ST_TblLayoutType.@fixed;
            table1.SetColumnWidth(0, 1200);
            table1.SetColumnWidth(1, 1200);
            table1.SetColumnWidth(2, 1200);

            XWPFTable table2 = doc.CreateTable(2, 3);
            var tblLayout2 = table2.GetCTTbl().tblPr.AddNewTblLayout();
            tblLayout2.type = ST_TblLayoutType.@fixed;
            table2.SetColumnWidth(0, 1500);
            table2.SetColumnWidth(1, 1500);
            table2.SetColumnWidth(2, 1500);
            table2.GetCTTbl().tblPr.AddNewTblPPr(); //tblPr.AddNewTblPPr is available since NPOI 2.5.5
            var tblpPr = table2.GetCTTbl().tblPr.tblpPr;
            tblpPr.leftFromText = 180;
            tblpPr.topFromText = 180;
            tblpPr.vertAnchor = ST_VAnchor.text;
            tblpPr.horzAnchor = ST_HAnchor.page;
            tblpPr.tblpY = "-840";  //this value is tricky, you have to calculate the vertical offset by the row number of the first table byyourself
            tblpPr.tblpX = "5800";

            using (FileStream fs = new FileStream("complexTable.docx", FileMode.Create))
            {
                doc.Write(fs);
            }
        }
    }
}
