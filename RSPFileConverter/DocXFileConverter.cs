using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace RSPFileConverter
{
    public static class DocXFileConverter
    {
        private const string fileExtension = "docx";

        /// <summary>
        /// Converts and download a pdf file from an Docx file
        /// </summary>
        /// <param name="filePath">The path for the source file to be converted</param>
        /// <param name="generatedPdfFileName">The final pdf file name</param>
        /// <returns>The path of the file converted</returns>
        public static string ToPdf(string filePath, string generatedPdfFileName) 
        {
            return ToPdf(filePath, generatedPdfFileName, new Dictionary<string, string>());
        }

        /// <summary>
        /// Converts and download a pdf file from an Docx file
        /// </summary>
        /// <param name="filePath">The path for the source file to be converted</param>
        /// <param name="generatedPdfFileName">The final pdf file name</param>
        /// <param name="keyValues">A collection of key values to be replaced inside the DocX file, it replaces each key with it's correspondent value</param>
        /// <returns>The path of the file converted</returns>
        public static string ToPdf(string filePath, string generatedPdfFileName, Dictionary<string, string> keyValues) 
        {
            string appDataPath = AppDomain.CurrentDomain.GetData("DataDirectory").ToString();
            string appDataTempPath = string.Format("{0}\\{1}\\", appDataPath, "Temp");
            if (!Directory.Exists(appDataTempPath))
            {
                Directory.CreateDirectory(appDataTempPath);
            }

            string docFileName = filePath.Substring(filePath.LastIndexOf('\\') + 1);            
            string tempDocFilePath = string.Format("{0}{1}_{2}.{3}", 
                appDataTempPath, docFileName, DateTime.Now.ToString("yyyy_MM_dd__HH_mm_ss"), fileExtension);
            string pdfFilePath = string.Format("{0}{1}_{2}.pdf",  
                appDataTempPath, generatedPdfFileName,  DateTime.Now.ToString("yyyy_MM_dd__HH_mm_ss"));

            File.Copy(filePath, tempDocFilePath, true);
            replaceValuesInDocFile(tempDocFilePath, keyValues);
            generatePdfFile(tempDocFilePath, pdfFilePath);
            DownloadFile(pdfFilePath, generatedPdfFileName);
            return pdfFilePath;
        }

        private static void replaceValuesInDocFile(string tempDocFilePath, Dictionary<string, string> keyValues)
        {
            using (WordprocessingDocument newWordDoc = WordprocessingDocument.Open(tempDocFilePath, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(newWordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                foreach (string key in keyValues.Keys)
                {
                    docText = docText.Replace(key, keyValues[key]);
                }

                using (StreamWriter sw = new StreamWriter(newWordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        private static void generatePdfFile(string filePath, string pdfFilePath)
        {
            Application appWord = new Microsoft.Office.Interop.Word.Application();
            Document wordDocument = appWord.Documents.Open(filePath);
            wordDocument.ExportAsFixedFormat(pdfFilePath, WdExportFormat.wdExportFormatPDF);
            wordDocument.Close();
        }

        private static void DownloadFile(string filePath, string fileName)
        {
            HttpResponse response = HttpContext.Current.Response;
            response.Buffer = true;
            response.Clear();
            response.AddHeader("Content-Disposition", "attachment;filename=" + fileName + ".pdf");
            response.ContentType = "application/pdf";
            response.BinaryWrite(File.ReadAllBytes(filePath));
            response.Flush();
        }
    }
}