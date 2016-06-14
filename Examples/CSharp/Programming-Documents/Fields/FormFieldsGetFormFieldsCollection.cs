using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FormFieldsGetFormFieldsCollection
    {
        public static void Run()
        {
            //ExStart:FormFieldsGetFormFieldsCollection
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            Document doc = new Document(dataDir + "FormFields.doc");
            FormFieldCollection formFields = doc.Range.FormFields;

            //ExEnd:FormFieldsGetFormFieldsCollection
            Console.WriteLine("\nDocument have " + formFields.Count.ToString() + " form fields.");
        }
    }
}
