using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using System.Text.RegularExpressions;
using System.Text;
namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Hyperlink
{
    class ReplaceHyperlinks
    {        
        public static void Run()
        {
            //ExStart:ReplaceHyperlinks
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithHyperlink();
            string NewUrl = @"http://www.aspose.com";
            string NewName = "Aspose - The .NET & Java Component Publisher";
            Document doc = new Document(dataDir + "ReplaceHyperlinks.doc");

            // Hyperlinks in a Word documents are fields.
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldHyperlink)
                {
                    FieldHyperlink hyperlink = (FieldHyperlink)field;

                    // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                    if (hyperlink.SubAddress != null)
                        continue;

                    hyperlink.Address = NewUrl;
                    hyperlink.Result = NewName;
                }
            }

            dataDir = dataDir + "ReplaceHyperlinks_out_.doc";
            doc.Save(dataDir);
            //ExEnd:ReplaceHyperlinks
            Console.WriteLine("\nHyperlinks replaced successfully.\nFile saved at " + dataDir);
        }
    }
}
