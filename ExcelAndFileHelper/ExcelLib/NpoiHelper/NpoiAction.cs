
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Text;

namespace ExcelLib.NpoiHelper
{
    public class NpoiAction
    {
        ISheet? excelSheet;
        public virtual DataSet ReadExcel(Stream stream)
        {
            List<string> rowList = new List<string>();
            stream.Position = 0;

            DataSet ds = new DataSet();
            try
            {
                using (XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream))
                {
                    ExcelToDataSet(xssWorkbook, ds, rowList);
                }
            }
            catch (Exception)
            {
                throw;
            }
            return ds;
        }

        public virtual byte[] CreateExcel(DataTable table)
        {
            byte[] result;
            try
            {
                using (var fs = new MemoryStream())
                {
                    table.TableName = string.IsNullOrEmpty(table.TableName) ? "Sheet1" : table.TableName;
                    IWorkbook workBook = new XSSFWorkbook();

                    CreateSheets(workBook, table);
                    workBook.Write(fs);
                    result = fs.ToArray();
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
            byte[] result;
            try
            {
                using (var ms = new MemoryStream())
                {
                    IWorkbook workBook = new XSSFWorkbook();
                    foreach (DataTable table in ds.Tables)
                    {
                        table.TableName = string.IsNullOrEmpty(table.TableName) ? "Sheet1" : table.TableName;
                        CreateSheets(workBook, table);
                    }
                    workBook.Write(ms);
                    result = ms.ToArray();
                }
            }
            catch (Exception)
            {

                throw;
            }
            return result;
        }

        public virtual byte[] ConvertExcelToCsv(Stream excelFile)
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                IWorkbook workbook = new XSSFWorkbook(excelFile);
                ISheet sheet =  workbook.GetSheetAt(0);
                for(int i =0;i<= sheet.LastRowNum;i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                
                    for(int j=0;j<row.LastCellNum;j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell !=null)
                        {
                            sb.Append(cell.ToString());
                        }
                        if(j< row.LastCellNum -1)
                        {
                            sb.Append(',');
                        }
                    }
                    sb.Append('\n');
                }
            }
            catch (Exception)
            {
                throw;
            }


            return Encoding.ASCII.GetBytes(sb.ToString());
        }

        public virtual string ReadExcelAsString(Stream stream)
        {
            List<string> rowList = new List<string>();
            stream.Position = 0;
            string sb = string.Empty;
            try
            {
                using (XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream))
                {
                    sb = ExcelToString(xssWorkbook);
                }
            }
            catch (Exception)
            {
                throw;
            }
            return sb;
        }
        private void ExcelToDataSet(XSSFWorkbook xssWorkbook,DataSet ds, List<string> rowList)
        {
            try
            {
                for (int k = 0; k < xssWorkbook.Count; k++)
            {
                
                excelSheet = xssWorkbook.GetSheetAt(k);
                DataTable dtTable = new DataTable();

                IRow headerRow = excelSheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;

                //Header part
                for (int j = 0; j < cellCount; j++)
                {
                    ICell cell = headerRow.GetCell(j);
                    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                    {
                        dtTable.Columns.Add(cell.ToString());
                        
                    }
                }

                //Row part
                for (int i = (excelSheet.FirstRowNum + 1); i <= excelSheet.LastRowNum; i++)
                {
                    IRow row = excelSheet.GetRow(i);
                    if (row == null) continue;
                    if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                            {
                                rowList.Add(row.GetCell(j).ToString());
                            }
                        }
                        
                    }

                    if (rowList.Count > 0)
                        dtTable.Rows.Add(rowList.ToArray());
                    rowList.Clear();

                }

                ds.Tables.Add(dtTable);
            }
            }
            catch (Exception)
            {
                throw;
            }
        }
        private string ExcelToString(XSSFWorkbook xssWorkbook)
        {
            StringBuilder? sb = new StringBuilder();
            try
            {
                for (int k = 0; k < xssWorkbook.Count; k++)
            {
                sb.Append($"Excel Sheet Name : {xssWorkbook.GetSheetName(k)}");
                sb.Append("----------------------------------------------- ");
                sb.Append('\n');

                excelSheet = xssWorkbook.GetSheetAt(k);
                

                IRow headerRow = excelSheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;

                //Header part
                for (int j = 0; j < cellCount; j++)
                {
                    ICell cell = headerRow.GetCell(j);
                    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                    {
                        
                        sb.Append(cell.ToString());

                        if (j < cellCount-1)
                            sb.Append(',');
                    }
                }

                sb.Append('\n');

                //Row part
                for (int i = (excelSheet.FirstRowNum + 1); i <= excelSheet.LastRowNum; i++)
                {
                    IRow row = excelSheet.GetRow(i);
                    if (row == null) continue;
                    if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                            {
                                

                                sb.Append(row.GetCell(j).ToString());

                                if (j < cellCount)
                                    sb.Append(",");
                            }
                        }
                    }

                    sb.Append('\n');
                }

                sb.Append('\n');
            }
            }
            catch (Exception)
            {
                throw;
            }
            return sb.ToString();
        }
        private void CreateSheets(IWorkbook workBook, DataTable table)
        {
            try
            {
                excelSheet = workBook.CreateSheet(table.TableName);

                List<string> columns = new List<string>();
                IRow row = excelSheet.CreateRow(0);
                int columnIndex = 0;
                int colLength = table.Columns.Count;

                foreach (DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);
                    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                    
                }

                int rowIndex = 1;
                foreach (DataRow dsRow in table.Rows)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (string col in columns)
                    {
                        row.CreateCell(cellIndex).SetCellValue(dsRow[col].ToString());
                        cellIndex++; 
                    }
                    rowIndex++;
                   
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
