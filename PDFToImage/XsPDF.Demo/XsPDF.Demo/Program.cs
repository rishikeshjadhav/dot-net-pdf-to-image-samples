using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XsPDF.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                PdfToImageDemo.ConvertPdfToImage();
                //BarcodeDemo.AddQRCodeToPDF();
                //ChartDemo.AddAreaChartToPDF();
                //CreatePDFDemo.CreateNewPDF();
                //DocumentProtectDemo.ProtectDocument();
                //ExtractDemo.ExtractTextFromPDF();
                //InsertToPDFDemo.CreateLongParagraphToPDF();
                //ProcessDemo.CombinePDFs();
                //MSChartDemo.InsertMSChartToPDF();
                //MSChartDemo.CreateMSRangeColumnChart();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
        }
    }
}
