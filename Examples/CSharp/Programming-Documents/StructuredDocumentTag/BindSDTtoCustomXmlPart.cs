using Aspose.Words.Markup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_StructuredDocumentTag
{
    class BindSDTtoCustomXmlPart
    {
        public static void Run()
        {
            // ExStart:BindSDTtoCustomXmlPart
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithStructuredDocumentTag();

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
    }
}
