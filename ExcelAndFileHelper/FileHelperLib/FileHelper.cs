using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileHelperLib
{
    public class FileHelper
    {
        public bool CompareTo(string file1,string file2)
        {
            int file1Byte;
            int file2Byte;
            FileStream fs1;
            FileStream fs2;

            if (!File.Exists(file1) || !File.Exists(file2))
                throw new FileNotFoundException("The specified file was not found.");

            if (string.IsNullOrEmpty(file1) || string.IsNullOrEmpty(file2))
                throw new ArgumentException("file1 and file2 are required");

            if (file1 == file2)
                return true;

            fs1 = new FileStream(file1,FileMode.Open,FileAccess.Read);
            fs2 = new FileStream(file2,FileMode.Open,FileAccess.Read);

            //check the file size. If they are not the same, the files are not the same.
            if(fs1.Length != fs2.Length)
            {
                //Close the file
                fs1.Close();
                fs2.Close();

                // Return false to indicate files are different
                return false;
            }

            //Read and compare a byte from each file until either a
            //non-matching set of bytes is found or until the end of file1 is reached.
            do
            {
                file1Byte = fs1.ReadByte();
                file2Byte = fs2.ReadByte();
            }while ((file1Byte == file2Byte) && (file1Byte != -1));

            //Close the files
            fs1.Close() ;
            fs2.Close() ;

            //Return the success of the comparison. "file1Byte" is
            //equal to file2Byte at this point only if the file are the same.
            return ((file1Byte - file2Byte) == 0);
        }

        public bool Move(string rootFolderPath,string destinationPath)
        {
            if (string.IsNullOrEmpty(rootFolderPath) && string.IsNullOrEmpty(destinationPath))
                throw new ArgumentException("rootFolderPath && destinationPath are required.");

            if(Directory.Exists(rootFolderPath) && Directory.Exists(destinationPath))
            {
                foreach (var file in new DirectoryInfo(rootFolderPath).GetFiles())
                {
                    file.MoveTo($@"{destinationPath}\{file.Name}");
                }

                return true;
            }
            else
            {
                throw new DirectoryNotFoundException("The specified directory was not found.");
            }
        }

        public bool Delete(string path)
        {
            if(string.IsNullOrEmpty(path))
                throw new ArgumentException("Path required");

            if (Directory.Exists(path)) 
            { 
                foreach (var file in new DirectoryInfo(path).GetFiles())
                {
                    file.Delete();
                }
                return true;
            }
            else
            {
                throw new DirectoryNotFoundException("The specified directory was not found.");
            }
        }

        public bool CopyTo(string rootPath,string destPath)
        {
            //Path should not be null
            if (string.IsNullOrEmpty(rootPath) && string.IsNullOrEmpty(destPath))
                throw new ArgumentException("RootPath and destPath are required.");

            if(Directory.Exists(rootPath) && Directory.Exists(rootPath))
            {
                foreach (var file in new DirectoryInfo(rootPath).GetFiles())
                {
                    file.CopyTo($@"{destPath}\{file.Name}");
                }
                return true;
            }
            else
            {
                throw new DirectoryNotFoundException("The specified Directory was not found");
            }
        }

        public bool IsFileExist(string path)
        {
            if(File.Exists(path)) return true; return false;
        }

        public bool IsDirectoryExist(string path) 
        { 
            if(File.Exists(path)) return true; return false;
        }

        public string CreateZip(string sourcePath,string destinationPath)
        {
            string returnPath = string.Empty;

            if (string.IsNullOrEmpty(sourcePath) && string.IsNullOrEmpty(destinationPath))
                throw new ArgumentNullException("sourcePath and destinationPath");

            string zipFileName = System.Guid.NewGuid().ToString() + ".zip";
            string zipFilePath = Path.Combine(destinationPath,zipFileName);
            try
            {
                ZipFile.CreateFromDirectory(sourcePath,zipFilePath);
                returnPath = zipFilePath;
            }catch (Exception ex)
            {
                throw;
            }
            return returnPath;
        }

        public string Unzip(string zipFilePath,string destinationPath)
        {
            string locationForExtract = string.Empty;
            if (string.IsNullOrEmpty(zipFilePath) && string.IsNullOrEmpty(destinationPath))
                throw new ArgumentNullException("zipFilePath and destinationPath are required.");

            try
            {
                string folderPath = Path.GetDirectoryName(zipFilePath);
                string folderToCreate = Path.GetFileNameWithoutExtension(zipFilePath);

                locationForExtract = Path.Combine(destinationPath, folderToCreate);

                ZipFile.ExtractToDirectory(zipFilePath,locationForExtract);
            }catch (Exception ex)
            {
                throw;
            }
            return locationForExtract;
        }
    }
}
