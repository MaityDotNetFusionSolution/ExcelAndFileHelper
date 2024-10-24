using ExcelLib.TextHelper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MimeMapping;

namespace ExcelApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TextHelperController : ControllerBase
    {
        private TextHelper _textHelper;
        public TextHelperController(TextHelper textHelper)
        {
             _textHelper = textHelper;
        }

        [HttpPost("ReadText", Name = "ReadTextFile")]
        public IActionResult ReadTextFile(IFormFile file)
        {
            if (file == null)
                return Ok("File Not Selected");

            string fileExtension = Path.GetExtension(file.FileName);
            if (fileExtension != ".txt" && fileExtension != ".txt")
                return Ok("File Not Selected");

            var data = _textHelper.ReadText(file.OpenReadStream());

            return Ok(data);
        }

        [HttpPost("CreateText",Name = "CreateTextFile")]
        public IActionResult CreateTextFile(string text)
        {
            string result = string.Empty;
            string FileName = "CreateText.txt";
            string contentType = MimeUtility.GetMimeMapping(FileName);

            var data = _textHelper.CreateTextFile(text);

            return File(data,contentType,FileName);
        }

        [HttpPost("ModifyText", Name = "ModifyTextFile")]
        public IActionResult ModifyTextFile(IFormFile file,string text)
        {
            string result = string.Empty;
            string FileName = "ModifyText.txt";
            string contentType = MimeUtility.GetMimeMapping(FileName);

            if (file == null || file.Length == 0)
                return Ok("File");

            string fileExtension = Path.GetExtension(file.FileName);
            if (fileExtension != ".txt" && fileExtension != ".txt")
                return Ok("File Not Selected");

            var data = _textHelper.ModifyTextFile(file.OpenReadStream(),text);

            return File(data, contentType, FileName);
        }
    }
}
