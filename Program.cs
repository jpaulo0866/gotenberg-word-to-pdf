using System;
using System.IO;
// using System.Linq;
// using Spire.Doc;
// using System.Drawing.Imaging;
// using System.Runtime.InteropServices;
// using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
// using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
// using System.Text;
// using System.Xml.Linq;
// using OpenXmlPowerTools;
// using OpenHtmlToPdf;
// using iText.Html2pdf;
// using iText.Kernel.Pdf;
using RestSharp;

namespace word_pdf
{
    class Program
    {
        private static string ORIGINAL_DOCX_LOCATION = "/home/joao_schmidt/test/original.docx";
        // private static string DOCX_LOCATION = "/home/joao_schmidt/test/documento.docx";
        private static string OTHER_DOCX_LOCATION = "/home/joao_schmidt/test/other.doc";

        private static string ANOTHER_DOCX_LOCATION = "/home/joao_schmidt/test/other2.docx";

        // private static string NEW_DOCX_LOCATION = "/home/joao_schmidt/test/documento-new.docx";
        // private static string NEW_DOCX_LOCATION = "/home/joao_schmidt/test/new.documento.docx";
        // private static string NEW_HTML_LOCATION = "/home/joao_schmidt/test/new.documento.html";
        // private static string HTML_IMAGE_LOCATION = "/home/joao_schmidt/test/images/";
        // private static string PDF_LOCATION = "/home/joao_schmidt/test/new.documento.pdf";
        private static string DATE_FORMAT = "yyyy-MM-dd hh:mm:ss";
        static void Main(string[] args)
        {
            //ClearPastFiles();
            //CopyOriginalFile();

            ConvertWithGotenberg(ORIGINAL_DOCX_LOCATION);
            ConvertWithGotenberg(OTHER_DOCX_LOCATION);
            ConvertWithGotenberg(ANOTHER_DOCX_LOCATION);

            //ConvertDocxToPdf();
            //EditOpenXmlDocument();
            //ConvertOpenXmlDocumentToHtml();
            //ConvertHtmlDocumentToPdf();
            //RenderHtmlDocumentToPdf();
            //ITextHtmlDocumentToPdf();
        }

        private static void ConvertWithGotenberg(string inputFile) {
            var client = new RestClient("http://localhost:3000");
            var request = new RestRequest("/convert/office", Method.POST);
            var newFileName = EditOpenXmlDocument(inputFile);

            request.AddFile(new FileInfo(newFileName).Name, newFileName);
            var response = client.Execute(request);

            File.WriteAllBytes(GetPdfExtensionedFileName(newFileName), response.RawBytes);
        }

        private static string EditOpenXmlDocument(string inputFile) {
            LogMessages("Editando o Docx");
            string newFileName = inputFile.Replace(".docx", "-new.docx").Replace(".doc", "-new.docx");

            Stream stream = File.Open(inputFile, FileMode.Open);
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true)) {

                string docText = null;
                using (StreamReader sr = new StreamReader(wordDocument.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex("«nome_proponente»");
                docText = regexText.Replace(docText, "Joao Schmidt");

                regexText = new Regex("«cidade»");
                docText = regexText.Replace(docText, "Porto alegre");

                regexText = new Regex("CIDADE");
                docText = regexText.Replace(docText, "Porto alegre");

                regexText = new Regex("«SEGURADO»");
                docText = regexText.Replace(docText, "Joao Schmidt");

                using (StreamWriter sw = new StreamWriter(wordDocument.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }

                LogMessages("Docx editado");
                LogMessages("Salvando novo Docx");
                var newFile = wordDocument.SaveAs(newFileName);
                newFile.Close();
                LogMessages("Novo Docx Salvo");
            }

            return newFileName;
        }

        private static string GetPdfExtensionedFileName(string inputFile) {
            return inputFile.Replace(".docx", ".pdf").Replace(".doc", ".pdf");
        }

        // private static void ConvertHtmlDocumentToPdf() {
        //     LogMessages("Iniciando conversão de HTML para PDF");
        //     string htmlContent = File.ReadAllText(NEW_HTML_LOCATION);
        //     LogMessages("HTML Renderizado");
        //     LogMessages("Abrindo novamente o Docx");
        //     Stream stream = File.Open(DOCX_LOCATION, FileMode.Open);
        //     LogMessages("Docx aberto");
        //     using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true)) {
        //         LogMessages("Convertendo para PDF");
        //         var pdfData = Pdf
        //                 .From(htmlContent)
        //                 .OfSize(PaperSize.A4)
        //                 .WithoutOutline()
        //                 .WithMargins(1.25.Centimeters())
        //                 .Portrait()
        //                 .Content();

        //         LogMessages("PDF Gerado");
        //         File.WriteAllBytes(PDF_LOCATION, pdfData);
        //         LogMessages("PDF Salvo");
        //     }

        // }

        // private static void RenderHtmlDocumentToPdf() {
        //     LogMessages("Iniciando conversão de HTML para PDF");
        //     string htmlContent = File.ReadAllText(NEW_HTML_LOCATION);
        //     LogMessages("HTML Renderizado");

        //     using (MemoryStream ms = new MemoryStream())
        //     {
        //         var pdf = TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerator.GeneratePdf(htmlContent, PdfSharp.PageSize.A4);
        //         LogMessages("PDF Gerado");
        //         pdf.Save(ms);
        //         File.WriteAllBytes(PDF_LOCATION, ms.ToArray());
        //         LogMessages("PDF Salvo");
        //     }

        // }

        // private static void ITextHtmlDocumentToPdf() {
        //     LogMessages("Iniciando conversão de HTML para PDF");
        //     string htmlContent = File.ReadAllText(NEW_HTML_LOCATION);
        //     LogMessages("HTML Renderizado");

        //     using (FileStream htmlSource = File.Open(NEW_HTML_LOCATION, FileMode.Open))
        //     using (FileStream pdfDest = File.Open(PDF_LOCATION, FileMode.OpenOrCreate))
        //     {
        //         ConverterProperties converterProperties = new ConverterProperties();
        //         iText.Html2pdf.HtmlConverter.ConvertToPdf(htmlSource, pdfDest, converterProperties);
        //     }

        // }

        // private static void ConvertOpenXmlDocumentToHtml() {

        //     LogMessages("Iniciando conversão de Docx para HTML");
        //     byte[] byteArray = File.ReadAllBytes(NEW_DOCX_LOCATION);
        //     using (MemoryStream memoryStream = new MemoryStream())
        //     {
        //         memoryStream.Write(byteArray, 0, byteArray.Length);
        //         using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
        //         {
        //             var imageDirectoryName = HTML_IMAGE_LOCATION;
        //             LogMessages($"Diretorio de imagens: {imageDirectoryName}");
        //             int imageCounter = 0;

        //             var pageTitle = "";
        //             var part = wDoc.CoreFilePropertiesPart;
        //             if (part != null)
        //             {
        //                 pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? "Documento";
        //             }

        //             // TODO: Determine max-width from size of content area.
        //             WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
        //             {
        //                 AdditionalCss = "body { margin: 5px auto; }",
        //                 PageTitle = pageTitle,
        //                 FabricateCssClasses = true,
        //                 CssClassPrefix = "pt-",
        //                 RestrictToSupportedLanguages = false,
        //                 RestrictToSupportedNumberingFormats = false,
        //                 ImageHandler = imageInfo =>
        //                 {
        //                     DirectoryInfo localDirInfo = new DirectoryInfo(imageDirectoryName);
        //                     if (!localDirInfo.Exists) {
        //                         localDirInfo.Create();
        //                     }

        //                     ++imageCounter;
        //                     string extension = imageInfo.ContentType.Split('/')[1].ToLower();

        //                     ImageFormat imageFormat = null;
        //                     if (extension == "png")
        //                         imageFormat = ImageFormat.Png;
        //                     else if (extension == "gif")
        //                         imageFormat = ImageFormat.Gif;
        //                     else if (extension == "bmp")
        //                         imageFormat = ImageFormat.Bmp;
        //                     else if (extension == "jpeg")
        //                         imageFormat = ImageFormat.Jpeg;
        //                     else if (extension == "tiff")
        //                     {
        //                         // Convert tiff to gif.
        //                         extension = "gif";
        //                         imageFormat = ImageFormat.Gif;
        //                     }
        //     private static void ClearPastFiles() {
        //     LogMessages("Deletando arquivos antigos");

        //     var files = new string[] { DOCX_LOCATION, NEW_DOCX_LOCATION, NEW_HTML_LOCATION, PDF_LOCATION };
        //     files.ToList().ForEach(item => {
        //         if (File.Exists(item)) {
        //             LogMessages($"Deletando o arquivo {item}");
        //             File.Delete(item);
        //         }
        //     });

        //     var folders = new string[] { HTML_IMAGE_LOCATION };
        //     folders.ToList().ForEach(item => {
        //         if (Directory.Exists(item)) {
        //             LogMessages($"Deletando a pasta {item}");
        //             Directory.EnumerateFiles(item).ToList().ForEach(file => {
        //                 File.Delete(file);
        //             });
        //             Directory.Delete(item);
        //         }
        //     });
        // }                else if (extension == "x-wmf")
        //                     {
        //                         extension = "wmf";
        //                         imageFormat = ImageFormat.Wmf;
        //                     }

        //                     // If the image format isn't one that we expect, ignore it,
        //                     // and don't return markup for the link.
        //                     if (imageFormat == null)
        //                         return null;

        //                     string imageFileName = imageDirectoryName + "/image" +
        //                         imageCounter.ToString() + "." + extension;
        //                     try
        //                     {
        //                         imageInfo.Bitmap.Save(imageFileName, imageFormat);
        //                     }
        //                     catch (System.Runtime.InteropServices.ExternalException)
        //                     {
        //                         return null;
        //                     }
        //                     string imageSource = localDirInfo.Name + "/image" +
        //                         imageCounter.ToString() + "." + extension;

        //                     string base64 = Base64.ConvertToBase64(imageFileName);
        //                     base64 = $"data:image/{extension};base64, {base64}";

        //                     XElement img = new XElement(Xhtml.img,
        //                         new XAttribute(NoNamespace.src, base64),
        //                         imageInfo.ImgStyleAttribute,
        //                         imageInfo.AltText != null ?
        //                             new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
        //                     LogMessages($"img: {img.Value}");
        //                     return img;
        //                 }
        //             };
        //             XElement htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

        //             // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
        //             // we are using HTML5.
        //             var html = new XDocument(
        //                 new XDocumentType("html", null, null, null),
        //                 htmlElement);

        //             // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
        //             // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
        //             // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
        //             // for detailed explanation.
        //             //
        //             // If you further transform the XML tree returned by ConvertToHtmlTransform, you
        //             // must do it correctly, or entities will not be serialized properly.

        //             var htmlString = html.ToString(SaveOptions.DisableFormatting);
        //             LogMessages("Salvando HTML");
        //             File.WriteAllText(NEW_HTML_LOCATION, htmlString, Encoding.UTF8);
        //             LogMessages("HTML Salvo");
        //         }
        //     }

        // }

        // private static void ConvertDocxToPdf() {
        //     LogMessages("Iniciando leitura do Word");
        //     Spire.Doc.Document docx = new Spire.Doc.Document();
        //     docx.LoadFromFile(DOCX_LOCATION);
        //     LogMessages("Word renderizado");
        //     docx.Replace("«nome_proponente»", "Joao", true, true);

        //     LogMessages("Salvando um novo Docx");
        //     docx.SaveToFile(NEW_DOCX_LOCATION);

        //     LogMessages("Salvando em PDF");
        //     docx.SaveToFile(PDF_LOCATION, FileFormat.PDF);
        // }

        // private static void ConvertDocxToHtml() {
        //     LogMessages("Iniciando leitura do Word");
        //     Spire.Doc.Document docx = new Spire.Doc.Document();
        //     docx.LoadFromFile(DOCX_LOCATION);
        //     LogMessages("Word renderizado");
        //     docx.Replace("«nome_proponente»", "Joao", true, true);

        //     LogMessages("Salvando um novo Docx");
        //     docx.SaveToFile(NEW_DOCX_LOCATION);

        //     LogMessages("Salvando em PDF");
        //     docx.SaveToFile(PDF_LOCATION, FileFormat.Html);
        // }

        // private static void ClearPastFiles() {
        //     LogMessages("Deletando arquivos antigos");

        //     var files = new string[] { DOCX_LOCATION, NEW_DOCX_LOCATION, NEW_HTML_LOCATION, PDF_LOCATION };
        //     files.ToList().ForEach(item => {
        //         if (File.Exists(item)) {
        //             LogMessages($"Deletando o arquivo {item}");
        //             File.Delete(item);
        //         }
        //     });

        //     var folders = new string[] { HTML_IMAGE_LOCATION };
        //     folders.ToList().ForEach(item => {
        //         if (Directory.Exists(item)) {
        //             LogMessages($"Deletando a pasta {item}");
        //             Directory.EnumerateFiles(item).ToList().ForEach(file => {
        //                 File.Delete(file);
        //             });
        //             Directory.Delete(item);
        //         }
        //     });
        // }

        // private static void CopyOriginalFile() {
        //     LogMessages("Copiando o arquivo Original");
        //     File.Copy(ORIGINAL_DOCX_LOCATION, DOCX_LOCATION);
        // }

        private static void LogMessages(string message) {
            Console.WriteLine($"{DateTime.Now.ToString(DATE_FORMAT)}::: {message}");
        }
    }
}
