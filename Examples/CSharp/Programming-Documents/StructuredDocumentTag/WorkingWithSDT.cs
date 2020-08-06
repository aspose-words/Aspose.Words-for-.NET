using Aspose.Words.Markup;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_StructuredDocumentTag
{
    public class WorkingWithSDT
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithStructuredDocumentTag();
            BindSDTtoCustomXmlPart(dataDir);
            ClearContentsControl(dataDir);
            SetContentControlColor(dataDir);
            SetContentControlStyle(dataDir);
            CreatingTableRepeatingSectionMappedToCustomXmlPart(dataDir);
            MultiSectionSDT(dataDir);
        }

        public static void SetContentControlColor(string dataDir)
        {
            // ExStart:SetContentControlColor

            Document doc = new Document(dataDir + "input.docx");
            StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            sdt.Color = Color.Red;

            dataDir = dataDir + "SetContentControlColor_out.docx";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:SetContentControlColor
            Console.WriteLine("\nSet the color of content control successfully.");
        }

        public static void ClearContentsControl(string dataDir)
        {
            // ExStart:ClearContentsControl

            Document doc = new Document(dataDir + "input.docx");
            StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            sdt.Clear();

            dataDir = dataDir + "ClearContentsControl_out.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:ClearContentsControl
            Console.WriteLine("\nClear the contents of content control successfully.");
        }

        public static void BindSDTtoCustomXmlPart(string dataDir)
        {
            // ExStart:BindSDTtoCustomXmlPart
            Document doc = new Document();
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), "<root><text>Hello, World!</text></root>");

            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
            doc.FirstSection.Body.AppendChild(sdt);

            sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/text[1]", "");

            dataDir = dataDir + "BindSDTtoCustomXmlPart_out.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:BindSDTtoCustomXmlPart
            Console.WriteLine("\nCreation of an XML part and binding a content control to it successfully.");
        }

        public static void SetContentControlStyle(string dataDir)
        {
            // ExStart:SetContentControlStyle
            Document doc = new Document(dataDir + "input.docx");
            StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            Style style = doc.Styles[StyleIdentifier.Quote];
            sdt.Style = style;

            dataDir = dataDir + "SetContentControlStyle_out.docx";
            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:SetContentControlStyle
            Console.WriteLine("\nSet the style of content control successfully.");
        }

        public static void CreatingTableRepeatingSectionMappedToCustomXmlPart(string dataDir)
        {
            // ExStart:CreatingTableRepeatingSectionMappedToCustomXmlPart
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

            doc.Save(dataDir + "Document.docx");
            // ExEnd:CreatingTableRepeatingSectionMappedToCustomXmlPart
            Console.WriteLine("\nCreation of a Table Repeating Section Mapped To a Custom Xml Part is successfull.");
        }

        public static void MultiSectionSDT(string dataDir)
        {
            // ExStart:MultiSectionSDT
            Document doc = new Document(dataDir + "input.docx");
            var tags = doc.GetChildNodes(NodeType.StructuredDocumentTagRangeStart, true);

            foreach (StructuredDocumentTagRangeStart tag in tags)
                Console.WriteLine(tag.Title);
            // ExEnd:MultiSectionSDT
        }
    }
}
