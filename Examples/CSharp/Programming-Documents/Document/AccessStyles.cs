
using System.IO;
using Aspose.Words;
using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class AccessStyles
    {
        public static void Run()
        {
            //ExStart:AccessStyles
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Load the template document.
            Document doc = new Document(dataDir + "TestFile.doc");
            // Get styles collection from document.
            StyleCollection styles = doc.Styles;
            string styleName = "";
            // Iterate through all the styles.
            foreach (Style style in styles)
            {
                if (styleName == "")
                {
                    styleName = style.Name;
                }
                else
                {
                    styleName = styleName + ", " + style.Name;
                }
            }
            //ExEnd:AccessStyles
            Console.WriteLine("\nDocument have following styles " + styleName);
        }        
    }
}
