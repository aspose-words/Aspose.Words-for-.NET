using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OpenXML_Word
{
    class Program
    {
        static void Main(string[] args)
        {
                 string docName="Test.docx";
                 WDDeleteHiddenText(docName);
        }
        public static void WDDeleteHiddenText(string docName)
        {
            // Given a document name, delete all the hidden text.
            const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(docName, true))
            {
                // Manage namespaces to perform XPath queries.
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("w", wordmlNamespace);

                // Get the document part from the package.
                // Load the XML in the document part into an XmlDocument instance.
                XmlDocument xdoc = new XmlDocument(nt);
                xdoc.Load(wdDoc.MainDocumentPart.GetStream());
                XmlNodeList hiddenNodes = xdoc.SelectNodes("//w:vanish", nsManager);
                foreach (System.Xml.XmlNode hiddenNode in hiddenNodes)
                {
                    XmlNode topNode = hiddenNode.ParentNode.ParentNode;
                    XmlNode topParentNode = topNode.ParentNode;
                    topParentNode.RemoveChild(topNode);
                    if (!(topParentNode.HasChildNodes))
                    {
                        topParentNode.ParentNode.RemoveChild(topParentNode);
                    }
                }

                // Save the document XML back to its document part.
                xdoc.Save(wdDoc.MainDocumentPart.GetStream(FileMode.Create, FileAccess.Write));
            }
        }

    }
}
