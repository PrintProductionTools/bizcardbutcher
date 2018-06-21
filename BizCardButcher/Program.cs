using BizCardButcher.Utilities;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace BizCardButcher
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourcePath = Path.GetFullPath(ConfigurationManager.AppSettings["sourcePath"])+@"\";
            Console.WriteLine("Verifying that following source path exists: " + sourcePath);
            if (!System.IO.Directory.Exists(sourcePath))
            {
                throw new FileNotFoundException(sourcePath + " does not exist on filesystem!");
            }
            string destPath = Path.GetFullPath(ConfigurationManager.AppSettings["destinationPath"])+@"\";
            Console.WriteLine("Creating destination path if it does not exist: " + destPath);
            Directory.CreateDirectory(destPath);
            DirectoryInfo dir = new DirectoryInfo(sourcePath);
            Console.WriteLine("Looking for PDFs in source path...");
            foreach (FileInfo file in dir.GetFiles("*.pdf"))
            {
                Console.WriteLine("Found " + file.Name);
                Console.WriteLine("Rotating PDF if in portrait layout...");
                string rotatedFile = PDFHelper.RotateIfNecessary(sourcePath + file.Name);
                Console.WriteLine("Cropping PDF...");
                string croppedFile = PDFHelper.Crop(sourcePath, destPath, rotatedFile);
                Console.WriteLine("Creating 10-up spread...");
                PDFHelper.CreateSpread(destPath, croppedFile);
                Console.WriteLine("Cleaning up intermediate files...");
                if (Path.GetFileName(rotatedFile) != file.Name) {
                    File.Delete(rotatedFile);
                }
                File.Delete(destPath + croppedFile);
                Console.WriteLine("Done processing " + file.Name);
            }
            Console.WriteLine("All done processing files.");
            Console.ReadLine();
        }
    }
}
