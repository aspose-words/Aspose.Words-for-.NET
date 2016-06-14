using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FormFieldsWorkWithProperties
    {
        public static void Run()
        {
            //ExStart:FormFieldsWorkWithProperties
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            Document doc = new Document(dataDir + "FormFields.doc");
            FormField formField = doc.Range.FormFields[3];

            if (formField.Type.Equals(FieldType.FieldFormTextInput))
                formField.Result = "My name is " + formField.Name;
            //ExEnd:FormFieldsWorkWithProperties            
        }
    }
}
