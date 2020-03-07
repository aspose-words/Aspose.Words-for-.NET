using Aspose.Words.Replacing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class IgnoreTextInsideFields
    {
        public static void Run()
        {
            // ExStart:IgnoreTextInsideFields
            // Create document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert field with text inside.
            builder.InsertField("INCLUDETEXT", "Text in field");

            Regex regex = new Regex("e");
            FindReplaceOptions options = new FindReplaceOptions();

            // Replace 'e' in document ignoring text inside field.
            options.IgnoreFields = true;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: \u0013INCLUDETEXT\u0014Text in field\u0015\f

            // Replace 'e' in document NOT ignoring text inside field.
            options.IgnoreFields = false;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: \u0013INCLUDETEXT\u0014T*xt in fi*ld\u0015\f
            // ExEnd:IgnoreTextInsideFields
        }
    }
}
