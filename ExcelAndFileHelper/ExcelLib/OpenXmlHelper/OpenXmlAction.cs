using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System.Data;
using System.Text;
#nullable disable

namespace ExcelLib.OpenXmlHelper
{
    public class OpenXmlAction
    {
        public virtual byte[] CreateExcel(DataTable table)
        {
            byte[]? result = null;
            try
            {
                using (var memory = new MemoryStream())
                {
                    //Create Spread Sheet
                    using (var doc = SpreadsheetDocument.Create(memory, SpreadsheetDocumentType.Workbook))
                    {
                        //Create WorkBookPart
                        WorkbookPart workbookPart = doc.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        //Create Worksheet
                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                        //Create Sheets collection
                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                        SheetDataCreation(workbookPart, worksheetPart, sheets, table, 1);

                        workbookPart.Workbook.Save();
                        memory.Position = 0;
                    }
                    result = memory.ToArray();
                }
            }
            catch (Exception)
            {

                throw;
            }
            return result;
        }

        public virtual byte[] CreateMultipleSheet(DataSet ds)
        {
            byte[]? result = null;
            try
            {
                using (var memory = new MemoryStream())
                {
                    ushort i = 1;
                    //Create Spread Sheet
                    using (var doc = SpreadsheetDocument.Create(memory, SpreadsheetDocumentType.Workbook))
                    {
                        //Create WorkBookPart
                        WorkbookPart workbookPart = doc.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        //Create Worksheet
                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                        //Create Sheets collection
                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                        //Multiple Sheet Creation 
                        foreach (DataTable table in ds.Tables)
                        {
                            SheetDataCreation(workbookPart, worksheetPart, sheets, table, i);
                            i++;
                        }

                        workbookPart.Workbook.Save();
                        memory.Position = 0;
                    }
                    result = memory.ToArray();
                }
            }
            catch (Exception)
            {

                throw;
            }
            return result;
        }

        public virtual DataSet ReadExcel(Stream stream)
        {
            string result = string.Empty;
            DataSet ds = new DataSet();
            try
            {

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, false))
                {
                    //Create workbook part
                    WorkbookPart workBookPart = doc.WorkbookPart;

                    //Create sheets
                    Sheets sheets = workBookPart.Workbook.GetFirstChild<Sheets>();

                    int i = 1;
                    foreach (Sheet sheet in sheets.OfType<Sheet>())
                    {
                        DataTable dt = new DataTable();
                        dt.TableName = $"Sheet{i}";

                        //statement to get the worksheet object by using the sheet id
                        Worksheet worksheet = ((WorksheetPart)workBookPart.GetPartById(sheet.Id)).Worksheet;

                        SheetData sheetData = worksheet.GetFirstChild<SheetData>();


                        for (int currentRow = 0; currentRow < sheetData.ChildElements.Count(); currentRow++)
                        {
                            List<string> rowList = new List<string>();
                            for (int curtCell = 0; curtCell < sheetData.ElementAt(currentRow).ChildElements.Count(); curtCell++)
                            {
                                DocumentFormat.OpenXml.Spreadsheet.Cell currentCell = (Cell)sheetData.ElementAt(currentRow).ChildElements.ElementAt(curtCell);
                                string currentCellValue = string.Empty;
                                if (currentCell.DataType != null)
                                {
                                    if (currentCell.DataType == CellValues.SharedString)
                                    {
                                        int id;
                                        if (Int32.TryParse(currentCell.InnerText, out id))
                                        {
                                            SharedStringItem item = workBookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                            if (item.Text != null)
                                            {
                                                //first row will provide the column name i.e CreateHeaderRow part
                                                if (currentRow == 0)
                                                {
                                                    dt.Columns.Add(item.Text.Text);
                                                }
                                                else
                                                {
                                                    rowList.Add(item.Text.Text);
                                                }
                                            }
                                            else if (item.InnerText != null)
                                            {
                                                currentCellValue = item.InnerText;
                                            }
                                            else if (item.InnerText != null)
                                            {
                                                currentCellValue = item.InnerXml;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (currentRow != 0)//reserved for column
                                    {
                                        rowList.Add(currentCell.InnerText);
                                    }
                                }
                            }
                            if (currentRow != 0)//reserved for column
                            {
                                dt.Rows.Add(rowList.ToArray());
                            }
                        }
                        ds.Tables.Add(dt);
                        i++;
                    }

                }
            }
            catch (Exception)
            {

                throw;
            }
            return ds;
        }

        public virtual string ReadExcelAsString(Stream stream)
        {
            DataSet ds = new DataSet();
            StringBuilder excelResult = new StringBuilder();
            try
            {
                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, false))
                {
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets sheetCollection = workbookPart.Workbook.GetFirstChild<Sheets>();


                    //using for each loop to get the sheet from the sheetCollection  
                    foreach (Sheet sheet in sheetCollection)
                    {
                        excelResult.AppendLine("Excel Sheet Name : " + sheet.Name);
                        excelResult.AppendLine("----------------------------------------------- ");
                        DataTable dt = new DataTable();

                        //statement to get the worksheet object by using the sheet id  
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;

                        SheetData sheetData = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                        foreach (Row currentRow in sheetData)
                        {
                            foreach (Cell currentCell in currentRow)
                            {
                                //statement to take the integer value  
                                string currentCellValue = string.Empty;
                                if (currentCell.DataType != null)
                                {
                                    if (currentCell.DataType == CellValues.SharedString)
                                    {
                                        int id;
                                        if (Int32.TryParse(currentCell.InnerText, out id))
                                        {
                                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                            if (item.Text != null)
                                            {
                                                //code to take the string value  
                                                excelResult.Append(item.Text.Text + " ");
                                            }
                                            else if (item.InnerText != null)
                                            {
                                                currentCellValue = item.InnerText;
                                            }
                                            else if (item.InnerXml != null)
                                            {
                                                currentCellValue = item.InnerXml;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    excelResult.Append(Convert.ToInt16(currentCell.InnerText) + " ");
                                }
                            }
                            excelResult.AppendLine();
                        }

                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return excelResult.ToString();
        }

        public virtual byte[] ConvertExcelToCsv(Stream excelFile)
        {
            byte[] bytes = null;
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(excelFile, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                var csvData = CsvString(workbookPart, sheetData);
                bytes = Encoding.ASCII.GetBytes(csvData);
            }
            return bytes;
        }

        private string CsvString(WorkbookPart workbookPart, SheetData sheetData)
        {
            StringBuilder sb = new StringBuilder();
            foreach (Row row in sheetData.Elements<Row>())
            {
                bool firstCell = true;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (!firstCell)
                    {
                        sb.Append(',');
                    }

                    string cellValue = GetCellValue(cell, workbookPart);

                    sb.Append(cellValue);

                    firstCell = false;
                }
                sb.Append('\n');
            }
            return sb.ToString();
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            string value = cell.CellValue?.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return workbookPart.SharedStringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
            }
            return value;
        }

        private void CreateHeaderRow(DataTable table, List<string> columns, Row headerRow, SheetData sheetData)
        {
            foreach (DataColumn column in table.Columns)
            {
                columns.Add(column.ColumnName);

                Cell cell = new Cell();
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(column.ColumnName);
                headerRow.AppendChild(cell);
            }

            sheetData.AppendChild(headerRow);
        }

        private void RowBaseOperation(DataTable table, List<string> columns, SheetData sheetData)
        {
            foreach (DataRow dsRow in table.Rows)
            {
                Row newRow = new Row();
                foreach (string col in columns)
                {
                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(dsRow[col].ToString());
                    newRow.AppendChild(cell);
                }

                sheetData.AppendChild(newRow);
            }
        }

        private void SheetDataCreation(WorkbookPart workbookPart, WorksheetPart worksheetPart, Sheets sheets, DataTable table, ushort i)
        {
            try
            {
                table.TableName = string.IsNullOrEmpty(table.TableName) ? $"Sheet{i}" : table.TableName;
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                //Create Sheet 1
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = i, Name = table.TableName };

                //Adding sheet to Sheet collection
                sheets.Append(sheet);


                Row headerRow = new Row();

                List<string> columns = new List<string>();

                CreateHeaderRow(table, columns, headerRow, sheetData);

                RowBaseOperation(table, columns, sheetData);
            }
            catch (Exception)
            {

                throw;
            }
        }
    }


}
