using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FormFieldsGetByName
    {
        public static void Run()
        {
            //ExStart:FormFieldsGetByName
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            Document doc = new Document(dataDir + "FormFields.doc");
            FormFieldCollection documentFormFields = doc.Range.FormFields;

            FormField formField1 = documentFormFields[3];
            FormField formField2 = documentFormFields["Text2"];
            //ExEnd:FormFieldsGetByName
            Console.WriteLine("\n" + formField2.Name + " field have following text " + formField2.GetText() + ".");
        }
    }
}
