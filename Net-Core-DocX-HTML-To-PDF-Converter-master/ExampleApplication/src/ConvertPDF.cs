using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using DocXToPdfConverter;
using DocXToPdfConverter.DocXToPdfHandlers;

namespace ExampleApplication.src
{
   public class ConvertPDF
    {
        

        public class Datum
        {
            public object parametro1 { get; set; }
            public object parametro2 { get; set; }
            public object parametro3 { get; set; }
            public object parametro4 { get; set; }
            public object parametro5 { get; set; }

            
        }

        public class Root
        {
            public string FileOutout { get; set; }
            public string FileInput { get; set; }
            public Datum Data { get; set; }
        }

        public static String ConvertAndPlaceholder(Root param)
        {
            try
            {
                string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                string locationOfLibreOfficeSoffice = Path.Combine(executableLocation, @"LibreOfficePortable\App\libreoffice\program\soffice.exe");

                string docxLocation = Path.Combine(executableLocation, param.FileInput);

                string htmlLocation = Path.Combine(executableLocation, "Test-HTML-page.html");


                var placeholders = new Placeholders
                {
                    NewLineTag = "<br/>",
                    TextPlaceholderStartTag = "##",
                    TextPlaceholderEndTag = "##",
                    TablePlaceholderStartTag = "==",
                    TablePlaceholderEndTag = "==",
                    ImagePlaceholderStartTag = "++",
                    ImagePlaceholderEndTag = "++",


                    TextPlaceholders = new Dictionary<string, string>
                {
                    {"Name", param.Data.parametro1.ToString()},
                    {"Street", param.Data.parametro2.ToString()},
                    {"City", param.Data.parametro3.ToString() },
                    {"InvoiceNo", param.Data.parametro4.ToString() },
                    {"Total", param.Data.parametro5.ToString() },
                    {"Date", "28 Jul 2019" },
                    {"Website", "www.smartinmedia.com" }
                },

                    HyperlinkPlaceholders = new Dictionary<string, HyperlinkElement>
                {
                    {"Website", new HyperlinkElement{ Link= "http://www.smartinmedia.com", Text="www.smartinmedia.com" } }
                },


                    TablePlaceholders = new List<Dictionary<string, string[]>>
                {
                    new Dictionary<string, string[]>()
                    {
                        {"Name", new string[]{ "Homer Simpson", "Mr. Burns", "Mr. Smithers" }},
                        {"Department", new string[]{ "Power Plant", "Administration", "Administration" }},
                        {"Responsibility", new string[]{ "Oversight", "CEO", "Assistant" }},
                        {"Telephone number", new string[]{ "888-234-2353", "888-295-8383", "888-848-2803" }}
                    },
                    new Dictionary<string, string[]>()
                    {
                        {"Qty", new string[]{ "2", "5", "7" }},
                        {"Product", new string[]{ "Software development", "Customization", "Travel expenses" }},
                        {"Price", new string[]{ "U$ 2,000", "U$ 1,000", "U$ 1,500" }},
                    }
                }
                };

                var productImage = StreamHandler.GetFileAsMemoryStream(Path.Combine(executableLocation, "ProductImage.jpg"));

                var qrImage = StreamHandler.GetFileAsMemoryStream(Path.Combine(executableLocation, "QRCode.PNG"));

                var productImageElement = new ImageElement() { Dpi = 96, MemStream = productImage };
                var qrImageElement = new ImageElement() { Dpi = 300, MemStream = qrImage };

                placeholders.ImagePlaceholders = new Dictionary<string, ImageElement>
                {
                    {"QRCode", qrImageElement },
                    {"ProductImage", productImageElement }
                };

                var test = new ReportGenerator(locationOfLibreOfficeSoffice);

                ////Convert from HTML to HTML
                //test.Convert(htmlLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-HTML-page-out.html"), placeholders);

                ////Convert from HTML to PDF
                //test.Convert(htmlLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-HTML-page-out.pdf"), placeholders);

                ////Convert from HTML to DOCX
                //test.Convert(htmlLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-HTML-page-out.docx"), placeholders);

                ////Convert from DOCX to DOCX
                //test.Convert(docxLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-Template-out.docx"), placeholders);

                ////Convert from DOCX to HTML
                //test.Convert(docxLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), "Test-Template-out.html"), placeholders);

                //Convert from DOCX to PDF
                test.Convert(docxLocation, Path.Combine(Path.GetDirectoryName(htmlLocation), param.FileOutout), placeholders);

                return "ok";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
            
        }
        public static String ConvertDocPdf(Root param)
        {
            try
            {
                string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                string locationOfLibreOfficeSoffice = Path.Combine(executableLocation, @"LibreOfficePortable\App\libreoffice\program\soffice.exe");

                string docxLocation = Path.Combine(executableLocation, param.FileInput);

                var test = new ReportGenerator(locationOfLibreOfficeSoffice);

                test.Convert(docxLocation, Path.Combine(Path.GetDirectoryName(docxLocation), param.FileOutout));

                return "ok";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
            

            
        }
        public static String ConvertingPDF(Root param)
        {

          String response =  (param.Data == null) ?ConvertDocPdf(param):ConvertAndPlaceholder(param);   
            
          return response;
        }
    }
}
