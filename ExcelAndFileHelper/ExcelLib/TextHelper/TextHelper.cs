using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLib.TextHelper
{
    public class TextHelper
    {
        public string ReadText(Stream stream)
        {
            string? result = string.Empty;

            using (StreamReader reader = new StreamReader(stream))
            {
                result =  reader.ReadToEnd();
                //result = reader.ReadLine();
            }

            return result;
        }

        public Stream CreateTextFile(string text)
        {
            MemoryStream stream = new MemoryStream();
            try
            {
                using (var writer = new StreamWriter(stream,leaveOpen:true))
                {
                    writer.WriteLine(text);
                    writer.Flush();
                }

                //Reset the stream position to beginning
                stream.Position = 0;
                return stream;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("An error occurred while modifying the text file",ex);
            }
        }

        public Stream ModifyTextFile(Stream stream,string text)
        {
            MemoryStream ms = new MemoryStream();
            try
            {

                using (var reader = new StreamReader(stream))
                {
                    string content =  reader.ReadToEnd();

                    //Modify the content
                    using (var writer = new StreamWriter(ms,leaveOpen:true))
                    {
                        writer.WriteLine(content);
                        writer.WriteLine(text);
                        writer.Flush();
                    }
                }

                //Reset the stream position to the beginning
                ms.Position = 0;
                return ms;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("An error occurred while modifying the text file", ex);
            }
            
        }
    }
}
