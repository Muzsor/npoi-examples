using NPOI.XSSF.Extractor;
using NPOI.XSSF.UserModel;
using System.IO;


using (var stream = File.OpenRead("text.xlsx"))
using (var xssWorkbook = new XSSFWorkbook(stream))
{
    var excelExtractor = new XSSFExcelExtractor(xssWorkbook)
    {
        IncludeCellComments = false,
        IncludeHeaderFooter = false
    };
    var stringCellValue = excelExtractor.Text;
    Console.Write(stringCellValue);
    Console.ReadLine();
}