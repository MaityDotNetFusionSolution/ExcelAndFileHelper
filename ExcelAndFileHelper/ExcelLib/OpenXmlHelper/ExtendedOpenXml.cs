#nullable disable

using System.Data;

namespace ExcelLib.OpenXmlHelper
{
    public class ExtendedOpenXml : OpenXmlAction
    {
        public override byte[] CreateExcel(DataTable table) => base.CreateExcel(table);
        public override byte[] CreateMultipleSheet(DataSet ds) => base.CreateMultipleSheet(ds);

        public override DataSet ReadExcel(Stream stream) => base.ReadExcel(stream);
        public override string ReadExcelAsString(Stream stream) => base.ReadExcelAsString(stream);
    }
}
