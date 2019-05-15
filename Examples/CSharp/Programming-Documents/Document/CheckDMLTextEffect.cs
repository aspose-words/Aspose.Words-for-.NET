using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CheckDMLTextEffect
    {
        public static void Run()
        {
            // ExStart:CheckDMLTextEffect
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            // Initialize document.
            Document doc = new Document(dataDir + "Document.doc");
            RunCollection runs = doc.FirstSection.Body.FirstParagraph.Runs;

            Font runFont = runs[0].Font;

            // One run might have several Dml text effects applied.
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Shadow));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Effect3D));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Reflection));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Outline));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Fill));
            // ExEnd:CheckDMLTextEffect
        }
    }
}
