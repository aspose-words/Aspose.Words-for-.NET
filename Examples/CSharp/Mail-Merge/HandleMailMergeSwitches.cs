using Aspose.Words.Fields;
using Aspose.Words.MailMerging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    public class HandleMailMergeSwitches
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting();
            // Open an existing document.
            Document doc = new Document(dataDir + "MailMergeSwitches.docx");

            doc.MailMerge.FieldMergingCallback = new MailMergeSwitches();

            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(
                new string[] { "HTML_Name" },
                new object[] { "James Bond" });

            dataDir = dataDir + "MergeSwitches_out.doc";
            doc.Save(dataDir);

            Console.WriteLine("\nSimple Mail merge performed with array data successfully.\nFile saved at " + dataDir);
        }
    }

    // ExStart:HandleMailMergeSwitches
    public class MailMergeSwitches : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
        {
            if (e.FieldName.StartsWith("HTML"))
            {
                if (e.Field.GetFieldCode().Contains("\\b"))
                {
                    FieldMergeField field = e.Field;

                    DocumentBuilder builder = new DocumentBuilder(e.Document);
                    builder.MoveToMergeField(e.DocumentFieldName, true, false);
                    builder.Write(field.TextBefore);
                    builder.InsertHtml(e.FieldValue.ToString());

                    e.Text = "";
                }
            }
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {

        }
    }
    // ExEnd:HandleMailMergeSwitches
}
