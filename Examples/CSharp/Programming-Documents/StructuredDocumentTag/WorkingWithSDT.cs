using Aspose.Words.Markup;
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
            SetContentControlStype(dataDir);
            BindSDTtoCustomXmlPart(dataDir);
            ClearContentsControl(dataDir);
            SetContentControlColor(dataDir);
            SetContentControlStype(dataDir);
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
            // ExStart:SetContentControlStype
            Document doc = new Document(dataDir + "input.docx");
            StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            Style style = doc.Styles[StyleIdentifier.Quote];
            sdt.Style = style;

            dataDir = dataDir + "SetContentControlStyle_out.docx";
            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:SetContentControlStype
            Console.WriteLine("\nSet the style of content control successfully.");
        }
    }
}
