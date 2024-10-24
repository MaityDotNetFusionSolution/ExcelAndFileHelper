using ExcelApi.Model;
using ExcelLib.NpoiHelper;

using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace ExcelApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class NpoiHelperController : ControllerBase
    {
        private ExtendedNpoi _npoi;
        private UserDetails _userDetails;

        public NpoiHelperController(ExtendedNpoi npoi, UserDetails userDetails)
        {
            _npoi = npoi;
            _userDetails = userDetails;
        }

        [HttpPost("ReadExcel", Name = "ReadExcelSheet")]
        public IActionResult ReadExcelSheet(IFormFile file)
        {
            FileValidationResp obj = new FileValidationResp(file);
            if (!obj.ValidateFileExcelFile().Status)
            {
                return Ok(obj.Message);
            }

            var result = _npoi.ReadExcel(file.OpenReadStream());
            var a = JsonConvert.SerializeObject(result, Formatting.Indented);
            return Ok(a);
        }

        [HttpPost("ReadExcelAsString")]
        public IActionResult ReadExcelAsString(IFormFile file)
        {

            FileValidationResp obj = new FileValidationResp(file);
            if (!obj.ValidateFileExcelFile().Status)
            {
                return Ok(obj.Message);
            }

            var result = _npoi.ReadExcelAsString(file.OpenReadStream());

            return Ok(result);
        }

        [HttpPost("CreateExcel", Name = "CreateExcelSheet")]
        public IActionResult CreateExcelSheet()
        {
            string FileName = "NpoiNewExcelFile.xlsx";
            string _contentType = MimeMapping.MimeUtility.GetMimeMapping(FileName);

            var table = _userDetails.ConvertModelToDataTable(_userDetails.GetEmployeeDummyData());
            var result = _npoi.CreateExcel(table);

            return File(result, _contentType, FileName);
        }

        [HttpPost("CreateMultipleExcelSheet")]
        public IActionResult CreateMultipleExcelSheet()
        {
            string FileName = "NpoiDummy.xlsx";
            string _contentType = MimeMapping.MimeUtility.GetMimeMapping(FileName);

            var ds = _userDetails.ConvertModelToDataSet(_userDetails.GetEmployeeDummyDataWithMultiList());
            var result = _npoi.CreateMultipleSheet(ds);

            return File(result, _contentType, FileName);
        }

        [HttpPost("ConvertExcelToCsv")]
        public IActionResult ConvertExcelToCsv(IFormFile file)
        {
            string FileName = "NewNpoiCsv.csv";
            string _contentType = MimeMapping.MimeUtility.GetMimeMapping(FileName);

            FileValidationResp obj = new FileValidationResp(file);
            if (!obj.ValidateFileExcelFile().Status)
            {
                return Ok(obj.Message);
            }

            var result = _npoi.ConvertExcelToCsv(file.OpenReadStream());

            return File(result, _contentType, FileName);
        }

    }
}
