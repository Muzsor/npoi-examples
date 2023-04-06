using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace CreateWorkbookFromDataSet
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var dataSet = CreateSampleData();
            using (IWorkbook workbook = new XSSFWorkbook())
            {

                for (var tableIndex = 0; tableIndex < dataSet.Tables.Count; tableIndex++)
                {
                    var dt = dataSet.Tables[tableIndex];
                    CreateSheetFromDataTable(workbook, tableIndex, dt);
                }

                FileStream sw = File.Create("File.xlsx");
                workbook.Write(sw, false);
                sw.Close();
            }
        }

        private static void CreateSheetFromDataTable(IWorkbook workbook, int dataTableIndex, DataTable dataTable)
        {
            var tableName = string.IsNullOrEmpty(dataTable.TableName) ? $"Sheet {dataTableIndex}" : dataTable.TableName;
            var sheet = (XSSFSheet)workbook.CreateSheet(tableName);
            var columnCount = dataTable.Columns.Count;
            var rowCount = dataTable.Rows.Count;

            // add column headers
            var row = sheet.CreateRow(0);
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                var col = dataTable.Columns[columnIndex];
                row.CreateCell(columnIndex).SetCellValue(col.ColumnName);
            }

            // add data rows
            for(var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var dataRow = dataTable.Rows[rowIndex];
                var sheetRow = sheet.CreateRow(rowIndex + 1);
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    sheetRow.CreateCell(columnIndex).SetCellValue(dataRow[columnIndex].ToString());
                }
            }
            
            // format the cell range as a table
            // note: if id, name, displayName are not set, Excel will not support the table
            // note: if id=0, Excel will not support the table
            XSSFTable xssfTable = sheet.CreateTable();
            CT_Table ctTable = xssfTable.GetCTTable();
            AreaReference myDataRange = new AreaReference(new CellReference(0, 0), new CellReference(rowCount, columnCount - 1));
            var tableId = uint.Parse((dataTableIndex + 1).ToString());
            ctTable.@ref = myDataRange.FormatAsString();
            ctTable.id = tableId;
            ctTable.name = $"Table{tableId}";
            ctTable.displayName = $"Table{tableId}";
            ctTable.tableStyleInfo = new CT_TableStyleInfo();
            ctTable.tableStyleInfo.name = "TableStyleMedium2"; // TableStyleMedium2 is one of XSSFBuiltinTableStyle
            ctTable.tableStyleInfo.showRowStripes = true;
            ctTable.tableColumns = new CT_TableColumns();
            ctTable.tableColumns.tableColumn = new List<CT_TableColumn>();
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                var col = dataTable.Columns[columnIndex];
                var colId = uint.Parse((columnIndex + 1).ToString());
                // note: if id=0, Excel will not support the table
                ctTable.tableColumns.tableColumn.Add(new CT_TableColumn() { id = colId, name = col.ColumnName });
            }

            // turn on filtering
            ctTable.autoFilter = new CT_AutoFilter();
            ctTable.autoFilter.@ref = myDataRange.FormatAsString();

            // auto size columns
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                sheet.AutoSizeColumn(columnIndex);
            }
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                // make room for the filter button and add a bit more
                var colWidth = sheet.GetColumnWidth(columnIndex);
                sheet.SetColumnWidth(columnIndex, colWidth + 1500);
            }
        }

        private static DataSet CreateSampleData()
        {
            var dataSet = new DataSet();

            // users
            var usersTable = new DataTable("Users");
            usersTable.Columns.Add("UserId");
            usersTable.Columns.Add("FirstName");
            usersTable.Columns.Add("LastName");
            usersTable.Rows.Add(1, "John", "Smith");
            usersTable.Rows.Add(2, "Sally", "Ride");
            usersTable.Rows.Add(3, "Jesse", "James");
            dataSet.Tables.Add(usersTable);

            // products
            var productsTable = new DataTable("Products");
            productsTable.Columns.Add("ProductId");
            productsTable.Columns.Add("Name");
            productsTable.Rows.Add(1, "Shirt");
            productsTable.Rows.Add(2, "Pants");
            productsTable.Rows.Add(3, "Shoes");
            dataSet.Tables.Add(productsTable);

            return dataSet;
        }
    }
}
