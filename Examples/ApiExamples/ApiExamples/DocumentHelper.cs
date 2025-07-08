﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System;
using System.IO;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    /// <summary>
    /// Functions for operations with documents and content
    /// </summary>
    internal class DocumentHelper : ApiExampleBase
    {
        /// <summary>
        /// Create simple document without run in the paragraph
        /// </summary>
        internal static Document CreateDocumentWithoutDummyText()
        {
            Document doc = new Document();

            //Remove the previous changes of the document
            doc.RemoveAllChildren();

            //Set the document author
            doc.BuiltInDocumentProperties.Author = "Test Author";

            //Create paragraph without run
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln();

            return doc;
        }

        /// <summary>
        /// Create new document with text
        /// </summary>
        internal static Document CreateDocumentFillWithDummyText()
        {
            Document doc = new Document();

            //Remove the previous changes of the document
            doc.RemoveAllChildren();

            //Set the document author
            doc.BuiltInDocumentProperties.Author = "Test Author";

            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Page ");
            builder.InsertField("PAGE", "");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES", "");

            //Insert new table with two rows and two cells
            InsertTable(builder);

            builder.Writeln("Hello World!");

            // Continued on page 2 of the document content
            builder.InsertBreak(BreakType.PageBreak);

            //Insert TOC entries
            InsertToc(builder);

            return doc;
        }

        internal static void FindTextInFile(string path, string expression)
        {
            using (var sr = new StreamReader(path))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();

                    if (string.IsNullOrEmpty(line)) continue;

                    if (line.Contains(expression))
                    {
                        Console.WriteLine(line);
                        Assert.Pass();
                    }
                    else
                    {
                        Assert.Fail();
                    }
                }
            }
        }

        /// <summary>
        /// Create new document template for reporting engine
        /// </summary>
        internal static Document CreateSimpleDocument(string templateText)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write(templateText);

            return doc;
        }

        /// <summary>
        /// Create new document with textbox shape and some query
        /// </summary>
        internal static Document CreateTemplateDocumentWithDrawObjects(string templateText, ShapeType shapeType)
        {
            Document doc = new Document();

            // Create textbox shape.
            Shape shape = new Shape(doc, shapeType);
            shape.Width = 431.5;
            shape.Height = 346.35;

            Paragraph paragraph = new Paragraph(doc);
            paragraph.AppendChild(new Run(doc, templateText));

            // Insert paragraph into the textbox.
            shape.AppendChild(paragraph);

            // Insert textbox into the document.
            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            return doc;
        }

        /// <summary>
        /// Compare word documents
        /// </summary>
        /// <param name="filePathDoc1">First document path</param>
        /// <param name="filePathDoc2">Second document path</param>
        /// <returns>Result of compare document</returns>
        internal static bool CompareDocs(string filePathDoc1, string filePathDoc2)
        {
            Document doc1 = new Document(filePathDoc1);
            Document doc2 = new Document(filePathDoc2);

            if (doc1.GetText() == doc2.GetText())
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Insert run into the current document
        /// </summary>
        /// <param name="doc">Current document</param>
        /// <param name="text">Custom text</param>
        /// <param name="paraIndex">Paragraph index</param>
        internal static Run InsertNewRun(Document doc, string text, int paraIndex)
        {
            Paragraph para = GetParagraph(doc, paraIndex);

            Run run = new Run(doc);
            run.Text = text;

            para.AppendChild(run);

            return run;
        }

        /// <summary>
        /// Insert text into the current document
        /// </summary>
        /// <param name="builder">Current document builder</param>
        /// <param name="textStrings">Custom text</param>
        internal static void InsertBuilderText(DocumentBuilder builder, string[] textStrings)
        {
            foreach (string textString in textStrings)
            {
                builder.Writeln(textString);
            }
        }

        /// <summary>
        /// Get paragraph text of the current document
        /// </summary>
        /// <param name="doc">Current document</param>
        /// <param name="paraIndex">Paragraph number from collection</param>
        internal static string GetParagraphText(Document doc, int paraIndex)
        {
            return doc.FirstSection.Body.Paragraphs[paraIndex].GetText();
        }

        /// <summary>
        /// Insert new table in the document
        /// </summary>
        /// <param name="builder">Current document builder</param>
        internal static Table InsertTable(DocumentBuilder builder)
        {
            //Start creating a new table
            Table table = builder.StartTable();

            //Insert Row 1 Cell 1
            builder.InsertCell();
            builder.Write("Date");

            //Set width to fit the table contents
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            //Insert Row 1 Cell 2
            builder.InsertCell();
            builder.Write(" ");

            builder.EndRow();

            //Insert Row 2 Cell 1
            builder.InsertCell();
            builder.Write("Author");

            //Insert Row 2 Cell 2
            builder.InsertCell();
            builder.Write(" ");

            builder.EndRow();

            builder.EndTable();

            return table;
        }

        /// <summary>
        /// Insert TOC entries in the document
        /// </summary>
        /// <param name="builder">
        /// The builder.
        /// </param>
        internal static void InsertToc(DocumentBuilder builder)
        {
            // Creating TOC entries
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 1.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;

            builder.Writeln("Heading 1.1.1.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading5;

            builder.Writeln("Heading 1.1.1.1.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading9;

            builder.Writeln("Heading 1.1.1.1.1.1.1.1.1");
        }

        /// <summary>
        /// Get section text of the current document
        /// </summary>
        /// <param name="doc">Current document</param>
        /// <param name="secIndex">Section number from collection</param>
        internal static string GetSectionText(Document doc, int secIndex)
        {
            return doc.Sections[secIndex].GetText();
        }

        /// <summary>
        /// Get paragraph of the current document
        /// </summary>
        /// <param name="doc">Current document</param>
        /// <param name="paraIndex">Paragraph number from collection</param>
        internal static Paragraph GetParagraph(Document doc, int paraIndex)
        {
            return doc.FirstSection.Body.Paragraphs[paraIndex];
        }

        /// <summary>
        /// Save the document to a stream, immediately re-open it and return the newly opened version
        /// </summary>
        /// <remarks>
        /// Used for testing how document features are preserved after saving/loading
        /// </remarks>
        /// <param name="doc">The document we wish to re-open</param>
        internal static Document SaveOpen(Document doc)
        {
            using (MemoryStream docStream = new MemoryStream())
            {
                doc.Save(docStream, new OoxmlSaveOptions(SaveFormat.Docx));
                return new Document(docStream);
            }
        }
    }
}
