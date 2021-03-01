// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class AddTable : TestUtil
    {
        [Test]
        public void AddTableFeature()
        {
            string[,] data = {{"Mike", "Amy"}, {"Mary", "Albert"}};

            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(ArtifactsDir + "Add Table - OpenXML.docx",
                    WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();

                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());

                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Create text in body - Create wordprocessing document"));

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

                table.AppendChild(props);

                for (var i = 0; i <= data.GetUpperBound(0); i++)
                {
                    var tr = new TableRow();
                    for (var j = 0; j <= data.GetUpperBound(1); j++)
                    {
                        var tc = new TableCell();
                        tc.Append(new Paragraph(new Run(new Text(data[i, j]))));

                        // Assume you want automatically sized columns.
                        tc.Append(new TableCellProperties(
                            new TableCellWidth {Type = TableWidthUnitValues.Auto}));

                        tr.Append(tc);
                    }

                    table.Append(tr);
                }

                mainPart.Document.Body.Append(table);
                mainPart.Document.Save();
            }
        }
    }
}

