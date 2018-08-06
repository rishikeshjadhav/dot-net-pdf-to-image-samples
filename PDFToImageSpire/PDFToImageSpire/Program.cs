using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Pdf;
using Spire.Pdf.Graphics;

namespace PDFToImageSpire
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //pdf file
                //String file = @"D:\PDF To Image Libs\201500663114__1_A301.pdf";
                //String file = @"D:\PDF To Image Libs\pdf-test.pdf";
                String file = "https://dodgepocstorage.blob.core.windows.net/cccc/201500663114__1_A301.pdf";

                //open pdf document
                PdfDocument doc = new PdfDocument();
                doc.LoadFromFile(file);
                //doc.loadfromstea
                FileStream from_stream = File.OpenRead("sample.pdf");

                //doc.LoadFromStream(from_stream);

                //save to images
                for (int i = 0; i < doc.Pages.Count; i++)
                {
                    String fileName = String.Format("Sample4-img-{0}.png", i);
                    using (Image image = doc.SaveAsImage(i))
                    {
                        image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
                        System.Diagnostics.Process.Start(fileName);
                    }
                }

                doc.Close();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
        }
    }
}
