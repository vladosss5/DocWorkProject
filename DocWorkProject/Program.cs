using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocWorkProject
{
    class Program
    {
        public static string InitialFilePath { get; private set; }

        static void Main(string[] args)
        {
            InitialFilePath = Environment.CurrentDirectory + "\\DocFileDirrect\\ConclusionTemplate.docx";

            string fileName = @"C:\users\public\documents\DocumentEx.docx";

            using (WordprocessingDocument doc = WordprocessingDocument.Open(InitialFilePath, true))
            {
                Console.WriteLine("Файл открыт");

                var mainPart = doc.MainDocumentPart;

                var docText = string.Empty;

                using (StreamReader sr = new StreamReader(mainPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }
            }

            // using (WordprocessingDocument myDocument =
            //        WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            // {
            //     // Add a main part.
            //     MainDocumentPart mainPart = myDocument.AddMainDocumentPart();
            //
            //     // Create the document structure.
            //     mainPart.Document = new Document();
            //     Body body = mainPart.Document.AppendChild(new Body());
            //     Paragraph para = body.AppendChild(new Paragraph());
            //     Run run = para.AppendChild(new Run());
            //
            //     // Add some text to the document.
            //     run.AppendChild(new Text("Hello, World!"));
            // }
            Console.WriteLine("The document has been created.\nPress a key.");
            Console.ReadKey();
        }
    }
}