using Aspose.Words.Fields;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ExtractTableOfContents
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            string fileName = "TOC.doc";
            Document doc = new Document(dataDir + fileName);

            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type.Equals(Aspose.Words.Fields.FieldType.FieldHyperlink))
                {
                    FieldHyperlink hyperlink = (FieldHyperlink)field;
                    if (hyperlink.SubAddress != null && hyperlink.SubAddress.StartsWith("_Toc"))
                    {
                        Paragraph tocItem = (Paragraph)field.Start.GetAncestor(NodeType.Paragraph);
                        Console.WriteLine(tocItem.ToString(SaveFormat.Text).Trim());
                        Console.WriteLine("------------------");
                        if (tocItem != null)
                        {
                            Bookmark bm = doc.Range.Bookmarks[hyperlink.SubAddress];
                            // Get the location this TOC Item is pointing to
                            Paragraph pointer = (Paragraph)bm.BookmarkStart.GetAncestor(NodeType.Paragraph);
                            Console.WriteLine(pointer.ToString(SaveFormat.Text));
                        }
                    } // End If
                }// End If
            }// End Foreach
        }
    }
}
