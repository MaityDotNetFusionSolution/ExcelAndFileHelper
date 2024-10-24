using FileHelperLib;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ExcelApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        private FileHelper _fileHelper;
        public FileController(FileHelper fileHelper)
        {
                _fileHelper = fileHelper;
        }

        [HttpPost("CompareTwoFiles",Name = "CompareTwoFiles")]
        public IActionResult CompareTwoFiles(string file1,string file2) 
            => Ok(_fileHelper.CompareTo(file1, file2) ? "Same File" : "Different file");

        [HttpPost("MoveFile")]
        public IActionResult MoveFile(string rootPath, string destinationPath)
            => Ok(_fileHelper.Move(rootPath, destinationPath) ? "File Move successfully" : "No file found");

        [HttpPost("DeleteFile")]
        public IActionResult DeleteFiles(string path, string file2)
             => Ok(_fileHelper.Delete(path) ? "File deleted" : "No file found");

        [HttpPost("CopyFile")]
        public IActionResult CopyFile(string rootPath, string destPath)
             => Ok(_fileHelper.CopyTo(rootPath, destPath) ? "File copied successfully" : "Failed to copy");

        [HttpPost("IsFileExist")]
        public IActionResult IsFileExist(string path)
             => Ok(_fileHelper.IsFileExist(path) ? "File present" : "File not present");

        [HttpPost("IsDirectoryExist")]
        public IActionResult IsDirectoryExist(string path)
             => Ok(_fileHelper.IsDirectoryExist(path) ? "Directory present" : "Directory not present");

        [HttpPost("CreateZip")]
        public IActionResult CreateZip(string sourcePath, string destinationPath)
             => Ok(_fileHelper.CreateZip(sourcePath, destinationPath));

        [HttpPost("Unzip")]
        public IActionResult Unzip(string zipFilePath, string destinationPath)
             => Ok(_fileHelper.Unzip(zipFilePath, destinationPath));
    }
}
