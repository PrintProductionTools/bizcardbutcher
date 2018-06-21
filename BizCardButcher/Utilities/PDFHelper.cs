using ImageMagick;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace BizCardButcher.Utilities
{
    internal class PDFHelper
    {
        // magic crop width/height for one-up card to place into multi-up template
        public static float CROP_WIDTH = ConversionHelper.InchesToDots(3.75f);

        public static float CROP_HEIGHT = ConversionHelper.InchesToDots(2.06f);

        // magic number to offset to align correctly on the slitter
        public static float NUDGE = ConversionHelper.MMToDots(1.8f);

        // magic numbers for margins and gutters for multi-up template
        public static float TOP_MARGIN = ConversionHelper.MMToDots(6.4f);

        public static float LEFT_MARGIN = ConversionHelper.MMToDots(12.7f) - NUDGE;
        public static float RIGHT_MARGIN = LEFT_MARGIN;
        public static float BOTTOM_MARGIN = ConversionHelper.MMToDots(6.2f) - (0.5f * NUDGE);
        public static float X_GUTTER = 2.22f * NUDGE;
        public static float Y_GUTTER = -0.15f * NUDGE;

        // magic number for one-up card dimensions on template
        public static float ONEUP_WIDTH = (1125f / 300 * 72) + NUDGE;

        public static float ONEUP_HEIGHT = (619f / 300 * 72) + NUDGE;

        public static string RotateIfNecessary(string pdfFile)
        {
            using (PdfReader reader = new PdfReader(pdfFile))
            {
                Rectangle pagesize = reader.GetPageSizeWithRotation(reader.GetPageN(1));
                if (pagesize.Width < pagesize.Height)
                {
                    for (int n = 1; n <= reader.NumberOfPages; n++)
                    {
                        PdfDictionary page = reader.GetPageN(n);
                        PdfNumber rotate = page.GetAsNumber(PdfName.ROTATE);
                        int rotation = rotate == null ? 90 : (rotate.IntValue + 90) % 360;
                        page.Put(PdfName.ROTATE, new PdfNumber(rotation));
                    }
                    string rotatedFileName = Path.GetDirectoryName(pdfFile)+@"\"+Path.GetFileNameWithoutExtension(pdfFile)+"_rotated.pdf";
                    using (FileStream fs = new FileStream(rotatedFileName, FileMode.OpenOrCreate,FileAccess.ReadWrite)) {
                        PdfStamper stamper = new PdfStamper(reader, fs);
                        stamper.Close();
                    }
                    return rotatedFileName;        
                } else return pdfFile;
            }
        }

        public static string Crop(string sourcePath, string destPath, string filePath)
        {
            string targetFile = Path.GetFileNameWithoutExtension(filePath) + "_crop.pdf";
            using (FileStream fs = new FileStream(destPath + targetFile, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                using (PdfReader reader = new PdfReader(filePath))
                {
                    float xOffset = (reader.GetPageSize(1).Width - CROP_WIDTH) / 2;
                    float yOffset = (reader.GetPageSize(1).Height - CROP_HEIGHT) / 2;
                    Rectangle cropSize = new Rectangle(CROP_WIDTH, CROP_HEIGHT);
                    using (Document doc = new Document(cropSize))
                    {
                        using (PdfWriter writer = PdfWriter.GetInstance(doc, fs))
                        {
                            doc.Open();
                            int pageCount = reader.NumberOfPages;
                            for (int currentPage = 0; currentPage < pageCount; currentPage++)
                            {
                                doc.NewPage();
                                PdfContentByte cb = writer.DirectContent;
                                PdfImportedPage page = writer.GetImportedPage(reader, currentPage + 1);
                                cb.AddTemplate(page, -xOffset, -yOffset);
                            }
                            doc.Close();
                        }
                    }
                }
            }
            return targetFile;
        }

        public static void CreateSpread(string destPath, string fileName)
        {
            string sourceFile = destPath + fileName;
            string targetFile = destPath + Path.GetFileNameWithoutExtension(fileName) + "_10up.pdf";

            using (FileStream fs = new FileStream(targetFile, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                using (Document doc = new Document(PageSize.LETTER))
                {
                    using (PdfWriter writer = PdfWriter.GetInstance(doc, fs))
                    {
                        using (PdfReader reader = new PdfReader(sourceFile))
                        {
                            doc.Open();
                            int pageCount = reader.NumberOfPages;
                            for (int currentPage = 0; currentPage < pageCount; currentPage++)
                            {
                                doc.NewPage();
                                for (int column = 0; column < 2; column++)
                                {
                                    for (int row = 0; row < 5; row++)
                                    {
                                        float x;
                                        float y = BOTTOM_MARGIN + (row * (ONEUP_HEIGHT + Y_GUTTER));
                                        if ((currentPage + 1) % 2 == 0)
                                        {
                                            x = doc.PageSize.Width + (3.5f * NUDGE) - RIGHT_MARGIN - ((column + 1) * (ONEUP_WIDTH + X_GUTTER));
                                        }
                                        else
                                        {
                                            float adjust = 3.5f * NUDGE;
                                            x = LEFT_MARGIN - NUDGE + (column * (ONEUP_WIDTH + X_GUTTER));
                                        }
                                        var tm = new System.Drawing.Drawing2D.Matrix();
                                        tm.Translate(x, y);
                                        var imp = writer.GetImportedPage(reader, (currentPage + 1));
                                        writer.DirectContent.AddTemplate(imp, tm);
                                    }
                                }
                            }
                            doc.Close();
                        }
                    }
                }
            }
        }
    }
}