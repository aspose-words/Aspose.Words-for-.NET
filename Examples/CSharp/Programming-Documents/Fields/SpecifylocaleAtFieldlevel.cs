using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class SpecifylocaleAtFieldlevel
    {
        public static void Run()
        { 
            // ExStart:SpecifylocaleAtFieldlevel
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            DocumentBuilder builder = new DocumentBuilder();
            Field field = builder.InsertField(FieldType.FieldDate, true);
            field.LocaleId = 1049;
			builder.Document.Save(dataDir + "SpecifylocaleAtFieldlevel_out.docx");
            // ExEnd:SpecifylocaleAtFieldlevel
        }
    }
}
