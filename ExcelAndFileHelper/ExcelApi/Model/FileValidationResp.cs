namespace ExcelApi.Model
{
#nullable disable
    public class FileValidationResp
    {
        public bool Status{ get; set; }
        public string? Message { get; set; }
        public string? Extension { get; set; }
        private IFormFile _file;
        public FileValidationResp Result { get; set; }
        public IFormFile file
        {
            get { return _file; }
            set { _file = value; }
        }


        public FileValidationResp(IFormFile file) {
            this.file = file; //?? throw new ArgumentNullException(nameof(file));

            if (file == null || file.Length == 0)
            {
                this.Message  = "File not selected";
                this.Status = false;
            }
            Extension = Path.GetExtension(file.FileName);
            if (Extension != ".xls" && Extension != ".xlsx")
            {
                this.Message = "Only excel file is allow";
                this.Status = false;
            }
            this.Status = true;
        }

        public FileValidationResp ValidateFileExcelFile()
        {
            return this;
        }
    }
}
