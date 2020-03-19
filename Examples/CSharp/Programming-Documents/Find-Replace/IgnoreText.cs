using Aspose.Words.Replacing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class IgnoreText
    {
        public static void Run()
        {
            IgnoreTextInsideFields();
            IgnoreTextInsideDeleteRevisions();
            IgnoreTextInsideInsertRevisions();
        }
        public static void IgnoreTextInsideFields()
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

        private static void IgnoreTextInsideDeleteRevisions()
        {
            // ExStart:IgnoreTextInsideDeleteRevisions
            // Create new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert non-revised text.
            builder.Writeln("Deleted");
            builder.Write("Text");

            // Remove first paragraph with tracking revisions.
            doc.StartTrackRevisions("author", DateTime.Now);
            doc.FirstSection.Body.FirstParagraph.Remove();
            doc.StopTrackRevisions();

            Regex regex = new Regex("e");
            FindReplaceOptions options = new FindReplaceOptions();

            // Replace 'e' in document ignoring deleted text.
            options.IgnoreDeleted = true;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: Deleted\rT*xt\f

            // Replace 'e' in document NOT ignoring deleted text.
            options.IgnoreDeleted = false;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: D*l*t*d\rT*xt\f
            // ExEnd:IgnoreTextInsideDeleteRevisions
        }

        private static void IgnoreTextInsideInsertRevisions()
        {
            // ExStart:IgnoreTextInsideInsertRevisions
            // Create new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text with tracking revisions.
            doc.StartTrackRevisions("author", DateTime.Now);
            builder.Writeln("Inserted");
            doc.StopTrackRevisions();

            // Insert non-revised text.
            builder.Write("Text");

            Regex regex = new Regex("e");
            FindReplaceOptions options = new FindReplaceOptions();

            // Replace 'e' in document ignoring inserted text.
            options.IgnoreInserted = true;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: Inserted\rT*xt\f

            // Replace 'e' in document NOT ignoring inserted text.
            options.IgnoreInserted = false;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: Ins*rt*d\rT*xt\f
            // ExEnd:IgnoreTextInsideInsertRevisions
        }
    }
}
