using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;

namespace _06._02_RemoveFormField
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Field field = builder.InsertField("PAGE");

            // Calling this method completely removes the field from the document.
            field.Remove();

            doc.Save("FormFieldTest.docx");
        }
    }
}
