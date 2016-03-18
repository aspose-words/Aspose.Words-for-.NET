using System;
using System.Collections.Generic;
using System.Text; 
using Aspose.Words;

namespace _05._04_UsingControlCharacters
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");

            // Enter a dummy field into the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD Field");

            // GetText will retrieve all field codes and special characters
            Console.WriteLine("GetText() Result: " + doc.GetText());

            string text = doc.GetText();
            text = text.Replace(ControlChar.Cr, ControlChar.CrLf);

            Console.WriteLine("Replaced text Result: " + text);
        }
    }
}
