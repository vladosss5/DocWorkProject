using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocWorkProject.Models;

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

            string[] baseTableWord =
            {
                "№", "Наименование", "Затратный подход", "Вес",
                "Сравнительный подход", "Вес", "Доходный подход", "Вес", "____ (вид) стоимость",
                "Учетная стоимость на дату оценки", ""
            };

            string[] childTableWord =
                { "Первоначальная стоимость по бухгалтерскому учету", "Остаточная стоимость по бухгалтерскому учету" };

            var countRow = Data.Estimates.Count + 3;
            var difference = baseTableWord.Length - childTableWord.Length;

            for (int i = 0; i < countRow; ++i)
            {
                var tr = new TableRow();

                for (int j = 0; j < baseTableWord.Length; ++j)
                {
                    var tc = new TableCell();

                    var paragraph = new Paragraph();
                    var run = new Run();

                    var runProperties = new RunProperties();
                    runProperties.Append(new RunFonts { Ascii = "Times New Roman" });
                    runProperties.Append(new FontSize { Val = "20" });

                    run.Append(runProperties);

                    var paragraphProperties = new ParagraphProperties();

                    paragraphProperties.Append(new Justification { Val = JustificationValues.Center });

                    paragraph.Append(paragraphProperties);
                    paragraph.Append(run);

                    var cellProps = new TableCellProperties();

                    cellProps.Append(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

                    if (i == 0)
                    {
                        if (j < baseTableWord.Length - 2)
                        {
                            run.Append(new Text(baseTableWord[j]));
                            cellProps.Append(new VerticalMerge { Val = MergedCellValues.Restart });
                        }
                        else if (j == baseTableWord.Length - 2)
                        {
                            run.Append(new Text(baseTableWord[j]));
                            cellProps.Append(new GridSpan { Val = 2 });
                            j++;
                        }
                    }
                    else if (i == 1)
                    {
                        if (j < baseTableWord.Length - 2)
                        {
                            cellProps.Append(new VerticalMerge { Val = MergedCellValues.Continue });
                        }
                        else
                        {
                            run.Append(new Text(childTableWord[j - difference]));
                        }
                    }
                    else if (i > 1 && i < countRow - 1)
                    {
                        switch (j)
                        {
                            case 0:
                                run.Append(new Text((i - 1).ToString()));
                                break;
                            case 1:
                                run.Append(new Text(Data.Estimates[i - 2].Name!));

                                var leftAlignProps = new ParagraphProperties();
                                leftAlignProps.Append(new Justification { Val = JustificationValues.Left });
                                paragraph.ParagraphProperties = leftAlignProps;
                                break;
                            case 9:
                                run.Append(new Text(Data.Estimates[i - 2].InitialCost + " руб."));
                                break;
                            case 10:
                                run.Append(new Text(Data.Estimates[i - 2].ResidualCost + " руб."));
                                break;
                        }
                    }
                    else if (i == countRow - 1)
                    {
                        if (j == 1)
                        {
                            run.Append(new Text("Итого"));
                            var leftAlignProps = new ParagraphProperties();
                            leftAlignProps.Append(new Justification { Val = JustificationValues.Left });
                            paragraph.ParagraphProperties = leftAlignProps;
                        }
                        else
                        {
                            run.Append(new Text(""));
                        }
                    }

                    tc.Append(cellProps); // Применить свойства ячейки
                    tc.Append(paragraph);
                    tr.Append(tc);
                }

                table.AppendChild(tr);
            }

            return table;
        }
    }
}