using PQScan.PDFToImage;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFToImagePQScan
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create an instance of PQScan.PDFToImage.PDFDocument object.
                PDFDocument pdfDoc = new PDFDocument();

                // Load PDF document from local file.
                pdfDoc.LoadPDF(@"D:\PDF To Image Libs\201500663114__1_A301.pdf");

                // Get the total page count.
                int count = pdfDoc.PageCount;

                for (int i = 0; i < count; i++)
                {
                    // Convert PDF page to image.
                    Bitmap jpgImage = pdfDoc.ToImage(i);

                    // Save image with jpg file type.
                    jpgImage.Save("output" + i + ".jpg", ImageFormat.Jpeg);
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
        }
    }
}
