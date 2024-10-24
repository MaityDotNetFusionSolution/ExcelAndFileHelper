using ExcelApi.Model;
using ExcelLib.CsvHelper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Net.Mime;

namespace ExcelApi.Controllers
{
    [Route("api/[controller]/[Action]")]
    [ApiController]
    public class CsvHelperController : ControllerBase
    {
        private CsvHelper _csv;private UserDetails _userDetails;
        public CsvHelperController(CsvHelper csv, UserDetails userDetails)
        {
            _csv = csv;
            _userDetails = userDetails;
        }

        [HttpPost("ReadCsv",Name = "ReadCsv")]
        public IActionResult ReadCsv(IFormFile file)
        {
            string result = string.Empty;
            if(file == null || file.Length == 0)
                    return Ok("File Not Selected");

            string fileExtension = Path.GetExtension(file.FileName);
            if (fileExtension != ".csv")
                return Ok("File Not Selected");

            result = JsonConvert.SerializeObject(_csv.ReadCsv(file.OpenReadStream()), Formatting.Indented);
            return Ok(result);
        }

        [HttpPost("CreateCsv", Name = "CreateCsvFile")]
        public IActionResult CreateCsvFile()
        {
            string FileName = "CreateCsvFile.csv";
            string contentType = MimeMapping.MimeUtility.GetMimeMapping(FileName);

            var dummyData = _userDetails.ConvertModelToDataTable(_userDetails.GetEmployeeDummyData());
            var data = _csv.CreateCsv(dummyData);

            
            return File(data, contentType, FileName);
        }
    }
}
