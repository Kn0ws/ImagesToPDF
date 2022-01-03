using System;
using System.Collections.Generic;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;

namespace ImageToPDF
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.Write("対象のフォルダをD&D: ");
            
            var inputMainFolder = Console.ReadLine();
            var mainFolder = inputMainFolder.Replace("\"", "");

            IEnumerable<string> subFolders = System.IO.Directory
                .EnumerateDirectories(
                    mainFolder,
                    "*",
                    System.IO.SearchOption.AllDirectories
                );
            Console.WriteLine();
            foreach (string subFolder in subFolders)
            {
                //Console.WriteLine(subFolder);
                var folderName = System.IO.Path.GetFileName(subFolder);
                Console.WriteLine(folderName + "内の画像ファイルを1つのPDFにまとめます・・・");
                generatePDF(subFolder, folderName);

            }
            Console.WriteLine("PDFの作成が完了しました。");
            Console.ReadLine();

        }


        public static void generatePDF(string subFolder, string folderName)
        {
            string[] files = System.IO.Directory
                .GetFiles(subFolder, "*.*")
                .OrderBy(f => f.Length)
                .Where(file => 
                file.ToLower().EndsWith("jpg")
                || file.ToLower().EndsWith("png")
                || file.ToLower().EndsWith("jpeg"))
                .ToArray();

            var document = new Document(iTextSharp.text.PageSize.A4, 25, 25, 25, 25);

            using (var stream = new FileStream(subFolder + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None))
            {
                PdfWriter.GetInstance(document, stream);
                document.Open();
                foreach (string file in files)
                {

                    using (var imageStream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {

                        var image = iTextSharp.text.Image.GetInstance(imageStream);

                        if (image.Height > image.Width)
                        {
                            document.SetPageSize(iTextSharp.text.PageSize.A4);
                            image.ScaleToFit(iTextSharp.text.PageSize.A4.Width - 25, iTextSharp.text.PageSize.A4.Height - 25);

                        }
                        else
                        {
                            document.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());
                            image.ScaleToFit(iTextSharp.text.PageSize.A4.Rotate().Width - 25, iTextSharp.text.PageSize.A4.Rotate().Height - 25);
                        }

                        image.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                        document.Add(image);
                    }
                }
                document.Close();
                Console.WriteLine(folderName + ".pdfを作成しました。");
                Console.WriteLine("");

            }
        }
    }
}
