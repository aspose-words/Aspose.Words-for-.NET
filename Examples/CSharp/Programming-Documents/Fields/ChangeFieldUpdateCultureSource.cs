using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class ChangeFieldUpdateCultureSource
    {
        public static void Run()
        {
            //ExStart:ChangeFieldUpdateCultureSource
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            // We will test this functionality creating a document with two fields with date formatting
            //ExStart:DocumentBuilderInsertField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert content with German locale.
            builder.Font.LocaleId = 1031;
            builder.InsertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
            builder.Write(" - ");
            builder.InsertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");
            //ExEnd:DocumentBuilderInsertField
            // Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
            // Set the culture used during field update to the culture used by the field.            
            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            doc.MailMerge.Execute(new string[] { "Date2" }, new object[] { new DateTime(2011, 1, 01) });
            dataDir = dataDir + "Field.ChangeFieldUpdateCultureSource_out_.doc";
            doc.Save(dataDir);
            //ExEnd:ChangeFieldUpdateCultureSource

            Console.WriteLine("\nCulture changed successfully used in formatting fields during update.\nFile saved at " + dataDir);
        }

    }
}
