using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class LoadAndSaveHtmlFormFieldasContentControlinDOCX
    {
        public static void Run()
        {
            // ExStart:LoadAndSaveHtmlFormFieldasContentControlinDOCX
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            HtmlLoadOptions lo = new HtmlLoadOptions();
            lo.PreferredControlType = HtmlControlType.StructuredDocumentTag;

            //Load the HTML document
            Document doc = new Document(dataDir + @"input.html", lo);

            //Save the HTML document into DOCX
            doc.Save(dataDir + "output.docx", SaveFormat.Docx);
            // ExEnd:LoadAndSaveHtmlFormFieldasContentControlinDOCX
            Console.WriteLine("\nHtml form fields are exported as content control successfully.");
        }
    }
}
