using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLib.NpoiHelper
{
    public class ExtendedNpoi : NpoiAction
    {
        public override DataSet ReadExcel(Stream stream) => base.ReadExcel(stream);
        public override byte[] CreateExcel(DataTable table) => base.CreateExcel(table);
        public override byte[] CreateMultipleSheet(DataSet ds) => base.CreateMultipleSheet(ds);
        public override byte[] ConvertExcelToCsv(Stream stream) => base.ConvertExcelToCsv(stream);
        public override string ReadExcelAsString(Stream stream) => base.ReadExcelAsString(stream);
        
    }
}
