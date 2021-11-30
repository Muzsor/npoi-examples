using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NPOI.Examples.XSSF.MeringCellsInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();

            var cell = sheet.CreateRow(1).CreateCell(1);
            cell.SetCellValue(new XSSFRichTextString("test1"));

            var style1=workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index2;
            style1.FillPattern = FillPattern.SolidForeground;

            cell.CellStyle = style1;

            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, 2));

            var cell2 = sheet.CreateRow(2).CreateCell(1);
            cell2.SetCellValue("test2");
            cell2.CellStyle = style1;

            sheet.AddMergedRegion(new CellRangeAddress(2, 2, 1, 2));

            var cra=new CellRangeAddress(2, 3, 4, 5);
            RegionUtil.SetBorderTop((int)BorderStyle.DashDot, cra, sheet);
            RegionUtil.SetBorderLeft((int)BorderStyle.DashDot, cra, sheet);
            RegionUtil.SetBorderRight((int)BorderStyle.DashDot, cra, sheet);
            RegionUtil.SetBorderBottom((int)BorderStyle.DashDot, cra, sheet);
            sheet.AddMergedRegion(cra);

            using (FileStream sw = File.Create("test.xlsx"))
            {
                workbook.Write(sw);
            }
        }
    }
}
