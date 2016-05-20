// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";

            string[,] data = new string[2, 2] { { "Mike", "Amy" }, { "Mary", "Albert" } };
            CreateWordprocessingDocument(FilePath + "Add Table - OpenXML.doc");
            AddTable(FilePath + "Add Table - OpenXML.doc", data);

        }

        private static void CreateWordprocessingDocument(string filepath)
        {
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
            }
        }

        private static void AddTable(string fileName, string[,] data)
        {
            using (var document = WordprocessingDocument.Open(fileName, true))
            {

                var doc = document.MainDocumentPart.Document;

                Table table = new Table();

                TableProperties props = new TableProperties(
                    new TableBorders(
                    new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    }));

                table.AppendChild<TableProperties>(props);

                for (var i = 0; i <= data.GetUpperBound(0); i++)
                {
                    var tr = new TableRow();
                    for (var j = 0; j <= data.GetUpperBound(1); j++)
                    {
                        var tc = new TableCell();
                        tc.Append(new Paragraph(new Run(new Text(data[i, j]))));

                        // Assume you want columns that are automatically sized.
                        tc.Append(new TableCellProperties(
                            new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                        tr.Append(tc);
                    }
                    table.Append(tr);
                }
                doc.Body.Append(table);
                doc.Save();
            }
        }
    }
}
