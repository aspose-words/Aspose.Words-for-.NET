﻿using System;
using System.Drawing;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Contents_Management
{
    internal class WorkingWithSdt : DocsExamplesBase
    {
        [Test]
        public void SdtCheckBox()
        {
            //ExStart:SdtCheckBox
            //GistId:089defec1b191de967e6099effeabda7
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
            builder.InsertNode(sdtCheckBox);
            
            doc.Save(ArtifactsDir + "WorkingWithSdt.SdtCheckBox.docx", SaveFormat.Docx);
            //ExEnd:SdtCheckBox
        }

        [Test]
        public void CurrentStateOfCheckBox()
        {
            //ExStart:CurrentStateOfCheckBox
            //GistId:089defec1b191de967e6099effeabda7
            Document doc = new Document(MyDir + "Structured document tags.docx");
            
            // Get the first content control from the document.
            StructuredDocumentTag sdtCheckBox =
                (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

            if (sdtCheckBox.SdtType == SdtType.Checkbox)
                sdtCheckBox.Checked = true;

            doc.Save(ArtifactsDir + "WorkingWithSdt.CurrentStateOfCheckBox.docx");
            //ExEnd:CurrentStateOfCheckBox
        }

        [Test]
        public void ModifySdt()
        {
            //ExStart:ModifySdt
            //GistId:089defec1b191de967e6099effeabda7
            Document doc = new Document(MyDir + "Structured document tags.docx");

            foreach (StructuredDocumentTag sdt in doc.GetChildNodes(NodeType.StructuredDocumentTag, true))
            {
                switch (sdt.SdtType)
                {
                    case SdtType.PlainText:
                    {
                        sdt.RemoveAllChildren();
                        Paragraph para = sdt.AppendChild(new Paragraph(doc)) as Paragraph;
                        Run run = new Run(doc, "new text goes here");
                        para.AppendChild(run);
                        break;
                    }
                    case SdtType.DropDownList:
                    {
                        SdtListItem secondItem = sdt.ListItems[2];
                        sdt.ListItems.SelectedValue = secondItem;
                        break;
                    }
                    case SdtType.Picture:
                    {
                        Shape shape = (Shape) sdt.GetChild(NodeType.Shape, 0, true);
                        if (shape.HasImage)
                        {
                            shape.ImageData.SetImage(ImagesDir + "Watermark.png");
                        }

                        break;
                    }
                }
            }
            
            doc.Save(ArtifactsDir + "WorkingWithSdt.ModifySdt.docx");
            //ExEnd:ModifySdt
        }

        [Test]
        public void SdtComboBox()
        {
            //ExStart:SdtComboBox
            //GistId:089defec1b191de967e6099effeabda7
            Document doc = new Document();

            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.ComboBox, MarkupLevel.Block);
            sdt.ListItems.Add(new SdtListItem("Choose an item", "-1"));
            sdt.ListItems.Add(new SdtListItem("Item 1", "1"));
            sdt.ListItems.Add(new SdtListItem("Item 2", "2"));
            doc.FirstSection.Body.AppendChild(sdt);

            doc.Save(ArtifactsDir + "WorkingWithSdt.SdtComboBox.docx");
            //ExEnd:SdtComboBox
        }

        [Test]
        public void SdtRichTextBox()
        {
            //ExStart:SdtRichTextBox
            //GistId:089defec1b191de967e6099effeabda7
            Document doc = new Document();

            StructuredDocumentTag sdtRichText = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);

            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc);
            run.Text = "Hello World";
            run.Font.Color = Color.Green;
            para.Runs.Add(run);
            sdtRichText.GetChildNodes(NodeType.Any, false).Add(para);
            doc.FirstSection.Body.AppendChild(sdtRichText);

            doc.Save(ArtifactsDir + "WorkingWithSdt.SdtRichTextBox.docx");
            //ExEnd:SdtRichTextBox
        }

        [Test]
        public void SdtColor()
        {
            //ExStart:SdtColor
            //GistId:089defec1b191de967e6099effeabda7
            Document doc = new Document(MyDir + "Structured document tags.docx");

            StructuredDocumentTag sdt = (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            sdt.Color = Color.Red;

            doc.Save(ArtifactsDir + "WorkingWithSdt.SdtColor.docx");
            //ExEnd:SdtColor
        }

        [Test]
        public void ClearSdt()
        {
            //ExStart:ClearSdt
            //GistId:089defec1b191de967e6099effeabda7
            Document doc = new Document(MyDir + "Structured document tags.docx");

            StructuredDocumentTag sdt = (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            sdt.Clear();

            doc.Save(ArtifactsDir + "WorkingWithSdt.ClearSdt.doc");
            //ExEnd:ClearSdt
        }

        [Test]
        public void BindSdtToCustomXmlPart()
        {
            //ExStart:BindSdtToCustomXmlPart
            //GistId:089defec1b191de967e6099effeabda7
            Document doc = new Document();
            CustomXmlPart xmlPart =
                doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), "<root><text>Hello, World!</text></root>");

            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
            doc.FirstSection.Body.AppendChild(sdt);

            sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/text[1]", "");

            doc.Save(ArtifactsDir + "WorkingWithSdt.BindSdtToCustomXmlPart.doc");
            //ExEnd:BindSdtToCustomXmlPart
        }

        [Test]
        public void SdtStyle()
        {
            //ExStart:SdtStyle
            //GistId:089defec1b191de967e6099effeabda7
            Document doc = new Document(MyDir + "Structured document tags.docx");

            StructuredDocumentTag sdt = (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            Style style = doc.Styles[StyleIdentifier.Quote];
            sdt.Style = style;

            doc.Save(ArtifactsDir + "WorkingWithSdt.SdtStyle.docx");
            //ExEnd:SdtStyle
        }

        [Test]
        public void RepeatingSectionMappedToCustomXmlPart()
        {
            //ExStart:RepeatingSectionMappedToCustomXmlPart
            //GistId:089defec1b191de967e6099effeabda7
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            CustomXmlPart xmlPart = doc.CustomXmlParts.Add("Books",
                "<books><book><title>Everyday Italian</title><author>Giada De Laurentiis</author></book>" +
                "<book><title>Harry Potter</title><author>J K. Rowling</author></book>" +
                "<book><title>Learning XML</title><author>Erik T. Ray</author></book></books>");

            Table table = builder.StartTable();

            builder.InsertCell();
            builder.Write("Title");

            builder.InsertCell();
            builder.Write("Author");

            builder.EndRow();
            builder.EndTable();

            StructuredDocumentTag repeatingSectionSdt =
                new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
            repeatingSectionSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book", "");
            table.AppendChild(repeatingSectionSdt);

            StructuredDocumentTag repeatingSectionItemSdt = 
                new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
            repeatingSectionSdt.AppendChild(repeatingSectionItemSdt);

            Row row = new Row(doc);
            repeatingSectionItemSdt.AppendChild(row);

            StructuredDocumentTag titleSdt =
                new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
            titleSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book[1]/title[1]", "");
            row.AppendChild(titleSdt);

            StructuredDocumentTag authorSdt =
                new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
            authorSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book[1]/author[1]", "");
            row.AppendChild(authorSdt);

            doc.Save(ArtifactsDir + "WorkingWithSdt.RepeatingSectionMappedToCustomXmlPart.docx");
            //ExEnd:RepeatingSectionMappedToCustomXmlPart
        }

        [Test]
        public void MultiSection()
        {
            //ExStart:MultiSectionSDT
            Document doc = new Document(MyDir + "Multi-section structured document tags.docx");

            NodeCollection tags = doc.GetChildNodes(NodeType.StructuredDocumentTagRangeStart, true);

            foreach (StructuredDocumentTagRangeStart tag in tags)
                Console.WriteLine(tag.Title);
            //ExEnd:MultiSectionSDT
        }

        [Test]
        public void SdtRangeStartXmlMapping()
        {
            //ExStart:SdtRangeStartXmlMapping
            //GistId:089defec1b191de967e6099effeabda7
            Document doc = new Document(MyDir + "Multi-section structured document tags.docx");

            // Construct an XML part that contains data and add it to the document's CustomXmlPart collection.
            string xmlPartId = Guid.NewGuid().ToString("B");
            string xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlPartContent);
            Console.WriteLine(Encoding.UTF8.GetString(xmlPart.Data));

            // Create a StructuredDocumentTag that will display the contents of our CustomXmlPart in the document.
            StructuredDocumentTagRangeStart sdtRangeStart = (StructuredDocumentTagRangeStart)doc.GetChild(NodeType.StructuredDocumentTagRangeStart, 0, true);

            // If we set a mapping for our StructuredDocumentTag,
            // it will only display a part of the CustomXmlPart that the XPath points to.
            // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart.
            sdtRangeStart.XmlMapping.SetMapping(xmlPart, "/root[1]/text[2]", null);

            doc.Save(ArtifactsDir + "WorkingWithSdt.SdtRangeStartXmlMapping.docx");
            //ExEnd:SdtRangeStartXmlMapping
        }
    }
}