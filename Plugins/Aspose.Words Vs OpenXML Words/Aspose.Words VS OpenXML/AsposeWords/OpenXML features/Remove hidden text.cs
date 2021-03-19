// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class RemoveHiddenText : TestUtil
    {
        [Test]
        public void RemoveHiddenTextFeature()
        {
            const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(MyDir + "Remove hidden text.docx", true))
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

                foreach (XmlNode hiddenNode in hiddenNodes)
                {
                    XmlNode topNode = hiddenNode.ParentNode.ParentNode;
                    XmlNode topParentNode = topNode.ParentNode;

                    topParentNode.RemoveChild(topNode);
                    if (!(topParentNode.HasChildNodes))
                        topParentNode.ParentNode.RemoveChild(topParentNode);
                }

                using (Stream stream = File.Create(ArtifactsDir + "Remove hidden text - OpenXML.docx"))
                {
                    xdoc.Save(stream);
                }
            }
        }
    }
}
