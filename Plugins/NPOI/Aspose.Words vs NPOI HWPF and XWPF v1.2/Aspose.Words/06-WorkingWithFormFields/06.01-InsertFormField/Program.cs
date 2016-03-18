using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace _06._01_InsertFormField
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Form Field
            string[] items = { "One", "Two", "Three" };
            builder.InsertComboBox("DropDown", items, 0);

            doc.Save("FormFieldTest.docx");
        }
    }
}
