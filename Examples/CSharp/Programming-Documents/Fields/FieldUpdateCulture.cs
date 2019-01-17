using Aspose.Words.Fields;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FieldUpdateCulture
    {
        public static void Run()
        {
            // ExStart:FieldUpdateCultureProvider
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            Document doc = new Document(dataDir + "FieldUpdateCultureProvider.docx");

            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            doc.FieldOptions.FieldUpdateCultureProvider = new FieldUpdateCultureProvider();

            dataDir = dataDir + "Field.FieldUpdateCultureProvider_out.pdf";
            doc.Save(dataDir);
            // ExEnd:FieldUpdateCultureProvider

            Console.WriteLine("\nFormat of Time field is set successfully.\nFile saved at " + dataDir);
        }
    }
}
