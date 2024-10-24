using DocumentFormat.OpenXml.EMMA;
using ExcelApi.Model;
using ExcelLib.OpenXmlHelper;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Data;

namespace ExcelApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class OpenXmlHelperController : ControllerBase
    {
        private ExtendedOpenXml _openXml;
        private UserDetails _userDetails;

        public OpenXmlHelperController(ExtendedOpenXml openXml, UserDetails userDetails)
        {
            _openXml = openXml;
            _userDetails = userDetails;
        }

        [HttpPost("ReadExcel")]
        public IActionResult ReadExcelSheet(IFormFile file)
        {

            FileValidationResp obj = new FileValidationResp(file);
            if (!obj.ValidateFileExcelFile().Status)
            {
                return Ok(obj.Message);
            }

            var result = _openXml.ReadExcel(file.OpenReadStream());
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

            var result = _openXml.ReadExcelAsString(file.OpenReadStream());

            return Ok(result);
        }

        [HttpPost("CreateExcel")]
        public IActionResult CreateExcelSheet()
        {
            string FileName = "OpenXmlNewExcelFile.xlsx";
            string _contentType = MimeMapping.MimeUtility.GetMimeMapping(FileName);

            var table = _userDetails.ConvertModelToDataTable(_userDetails.GetEmployeeDummyData());
            var result = _openXml.CreateExcel(table);

            return File(result, _contentType, FileName);
        }

        [HttpPost("CreateMultipleExcelSheet")]
        public IActionResult CreateMultipleExcelSheet()
        {
            string FileName = "OpenXmlDummy.xlsx";
            string _contentType = MimeMapping.MimeUtility.GetMimeMapping(FileName);

            var ds = _userDetails.ConvertModelToDataSet(_userDetails.GetEmployeeDummyDataWithMultiList());
            var result = _openXml.CreateMultipleSheet(ds);

            return File(result, _contentType, FileName);
        }

        [HttpPost("ConvertExcelToCsv")]
        public IActionResult ConvertExcelToCsv(IFormFile file)
        {
            string FileName = "NewOpenXmlCsv.csv";
            string _contentType = MimeMapping.MimeUtility.GetMimeMapping(FileName);

            FileValidationResp obj = new FileValidationResp(file);
            if (!obj.ValidateFileExcelFile().Status)
            {
                return Ok(obj.Message);
            }

            var result = _openXml.ConvertExcelToCsv(file.OpenReadStream());

            return File(result, _contentType, FileName);
        }
    }
}
