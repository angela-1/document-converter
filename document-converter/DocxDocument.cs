using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MSWord = Microsoft.Office.Interop.Word;

namespace document_converter
{
    class DocxDocument
    {
        private static string GetDestFilename(string filePath, OutputFormat formatType)
        {
            string newFile = Path.GetFileNameWithoutExtension(filePath) + "." + GetExtension(formatType);
            string dest = Path.GetDirectoryName(filePath);
            return Path.Combine(dest, newFile);
        }

        private static string GetExtension(OutputFormat formatType)
        {
            Dictionary<OutputFormat, string> dict = new Dictionary<OutputFormat, string>
            {
                { OutputFormat.PDF, "pdf" },
                { OutputFormat.DOCX, "docx" },
                { OutputFormat.TXT, "txt" }
            };
            return dict[formatType];
        }
        public static void Convert(string filePath, OutputFormat formatType)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine("文件" + filePath + "不存在");
                return;
            }

            MSWord.Application app = new MSWord.Application
            {
                Visible = false
            };

            MSWord.Document doc = app.Documents.Open(filePath);
            string dest = GetDestFilename(filePath, formatType);
            switch (formatType)
            {
                case OutputFormat.PDF:
                    doc.ExportAsFixedFormat(dest, MSWord.WdExportFormat.wdExportFormatPDF,
                        CreateBookmarks: MSWord.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks);
                    break;
                case OutputFormat.DOCX:
                    doc.SaveAs2(dest, MSWord.WdSaveFormat.wdFormatDocumentDefault);
                    break;
                case OutputFormat.TXT:
                    doc.SaveAs2(dest, MSWord.WdSaveFormat.wdFormatText);
                    break;
                default:
                    break;
            }

            doc.Close();
            app.Quit();
        }
    }
}
