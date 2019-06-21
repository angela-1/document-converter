using McMaster.Extensions.CommandLineUtils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace document_converter
{
    class Program
    {
        static int Main(string[] args)
        {
            return CommandLineApplication.Execute<Program>(args);
        }

        [Argument(0, Description = "Full path of docx file or folder")]
        public string FilePath { get; }

        [Option(ShortName = "f", LongName = "format", Description = "Output format, pdf, docx or txt")]
        public string Format { get; }

        private static List<string> FilterDocx(List<string> filepaths)
        {
            List<string> result = new List<string>();
            foreach (var item in filepaths)
            {
                String ext = Path.GetExtension(item);
                if (ext == ".docx" || ext == ".doc")
                {
                    result.Add(item);
                }
            }
            return result;
        }

        private static OutputFormat GetFormatType(string format)
        {
            OutputFormat formatType = OutputFormat.PDF;
            switch (format)
            {
                case "pdf":
                    formatType = OutputFormat.PDF;
                    break;
                case "docx":
                    formatType = OutputFormat.DOCX;
                    break;
                case "txt":
                    formatType = OutputFormat.TXT;
                    break;
                default:
                    break;
            }
            return formatType;
        }

        private void OnExecute()
        {
            bool v = FilePath == null;
            if (v)
            {
                Console.WriteLine("No input, please input file or folder");
                return;
            }
            bool isFolder = Directory.Exists(FilePath);
            string formatString = Format ?? "docx";
            OutputFormat outputFormat = GetFormatType(formatString);
            if (isFolder)
            {
                string[] filePaths = Directory.GetFiles(FilePath);
                List<string> docxFiles = FilterDocx(new List<string>(filePaths));
                foreach (var docx in docxFiles)
                {
                    DocxDocument.Convert(docx, outputFormat);
                    Console.WriteLine("Done");
                }
            }
            else
            {
                string ext = Path.GetExtension(FilePath);
                if (ext != ".docx" && ext != ".doc")
                {
                    Console.WriteLine($"Not support {ext} format");
                    return;
                }
                DocxDocument.Convert(FilePath, outputFormat);
                Console.WriteLine("Done");
            }
        }

    }
}
