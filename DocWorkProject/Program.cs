using System.Net.Mime;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocWorkProject.Models;
using Array = DocumentFormat.OpenXml.Office2019.Excel.RichData2.Array;
using DTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using DTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using DText = DocumentFormat.OpenXml.Wordprocessing.Text;
using DRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;

namespace DocWorkProject
{
    class Program
    {
        public static string InitialFilePath { get; private set; }
        public static WordprocessingDocument Document { get; protected set; }
        public static WordDataModel Data { get; protected set; }

        static void Main(string[] args)
        {
            InitialFilePath = Environment.CurrentDirectory + "\\DocFileDirrect\\ConclusionTemplate.docx";

            Data = new WordDataModel()
            {
                Id = Guid.NewGuid().ToString(),
                DateCompilation = DateOnly.FromDateTime(DateTime.Now),
                Number = "26354-345",
                Customer = "ЗАКАЗЧИК БИБА",
                Appraiser = "ОЦЕНЩИК БОБА",
                TypeCost = "КАКОЙТО ТИП СТОИМОСТИ",
                PurposeAssessment = "КАКАЯТО ЦЕЛЬ ОЦЕНКИ",
                DateAssessment = DateOnly.FromDateTime(DateTime.Now),
                DateCompilationReport = DateOnly.FromDateTime(DateTime.Now),
                Estimates = new List<Estimate>
                {
                    new()
                    {
                        Id = Guid.NewGuid().ToString(),
                        Name = "Какое-то имя",
                        InitialCost = 9000,
                        ResidualCost = 5000
                    }
                }
            };

            ReplaceTagsInWordDocument();
        }

        public static void ReplaceTagsInWordDocument()
        {
            using (var document = WordprocessingDocument.Open(InitialFilePath, true))
            {
                Document = document;
                var mainPart = Document.MainDocumentPart;

                var tags = mainPart.Document.Body.Descendants<Text>()
                    .Where(r => r.Text == "DateCompilation" ||
                                r.Text == "Number" ||
                                r.Text == "EvaluationResultsTable")
                    .ToList();

                foreach (var tag in tags)
                {
                    ReplaceTag(tag);
                }

                mainPart.Document.Save();
            }
        }

        private static void ReplaceTag(Text tag)
        {
            var body = Document.MainDocumentPart.Document.Body;

            switch (tag.Text)
            {
                case "DateCompilation":
                    tag.Text = Data.DateCompilation.ToString("dd.MM.yyyy");
                    break;

                case "Number":
                    tag.Text = Data.Number!;
                    break;

                case "EvaluationResultsTable":
                    var parentParagraph = tag.Parent.Parent as Paragraph;

                    if (parentParagraph != null)
                    {
                        var table = CreateTable();
                        var previousElement = parentParagraph.PreviousSibling();

                        body.RemoveChild(parentParagraph);

                        if (previousElement != null)
                        {
                            body.InsertAfter(table, previousElement);
                        }
                        else
                        {
                            body.InsertAt(table, 0);
                        }
                    }

                    break;
            }
        }


        public static Table CreateTable()
        {
            var table = new Table();

            var tableProperties = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 }
                )
            );
            table.AppendChild(tableProperties);
            
            string[] tableWord =  { "№", "Наименование", "Затратный подход", "Вес",
                "Сравнительный подход", "Вес", "Доходный подход",  "Вес", "____ (вид) стоимость",
                "Первоначальная стоимость по бухгалтерскому учету", "Остаточная стоимость, руб."};

            var countRow = Data.Estimates.Count + 2;

            for (int i = 0; i < countRow; ++i)
            {
                var tr = new TableRow();

                for (int j = 0; j < tableWord.Length; ++j)
                {
                    var tc = new TableCell();

                    if (i == 0)
                        tc.Append(new Paragraph(new Run(new Text(tableWord[j]))));

                    if (i == tableWord.Length)
                    {
                        
                    }

                    if (i <= countRow - 2 && i > 0)
                    {
                        switch (j)
                        {
                            case 0:
                                tc.Append(new Paragraph(new Run(new Text(i.ToString()))));
                                break;
                            case 1:
                                tc.Append(new Paragraph(new Run(new Text(Data.Estimates[i - 1].Name!))));
                                break;
                            case 9:
                                tc.Append(new Paragraph(new Run(new Text(Data.Estimates[i - 1].InitialCost.ToString()))));
                                break;
                            case 10:
                                tc.Append(new Paragraph(new Run(new Text(Data.Estimates[i - 1].ResidualCost.ToString()))));
                                break;
                        }
                    }

                    if (i == countRow - 1 && j == 1)
                        tc.Append(new Paragraph(new Run(new Text("Итого"))));

                    tr.Append(tc);
                }

                table.AppendChild(tr);
            }

            return table;
        }
    }
}