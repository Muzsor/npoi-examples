using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using (IWorkbook workbook = new XSSFWorkbook())
{
    ISheet sheet = workbook.CreateSheet("Sheet1");
    //increase the width of Column A
    sheet.SetColumnWidth(0, 5000);
    //create the format instance
    IDataFormat format = workbook.CreateDataFormat();

    
    //Chinese date string
    ICell cell7 = sheet.CreateRow(6).CreateCell(0);
    SetValueAndFormat(workbook, cell7, new DateOnly(2004, 5, 6), format.GetFormat("yyyy年m月d日"));

    using (FileStream sw = File.Create("test.xlsx"))
    {
        workbook.Write(sw, false);
    }
}

static void SetValueAndFormat(IWorkbook workbook, ICell cell, DateOnly value, short formatId)
{
    //set value for the cell
    cell.SetCellValue(value);

    ICellStyle cellStyle = workbook.CreateCellStyle();
    cellStyle.DataFormat = formatId;
    cell.CellStyle = cellStyle;
}