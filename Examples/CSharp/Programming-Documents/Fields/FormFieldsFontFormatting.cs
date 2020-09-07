using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FormFieldsFontFormatting
    {
        public static void Run()
        {
            // ExStart:FormFieldsFontFormatting
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            Document doc = new Document(dataDir + "Document.doc");
            doc.Range.FormFields[0].Font.Size = 20;
            doc.Range.FormFields[0].Font.Color = Color.Red;
            doc.Save(dataDir + "Document_out.doc");
            // ExEnd:FormFieldsFontFormatting
        }
    }
}
