
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Settings;
namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class SetViewOption
    {
        public static void Run()
        {
            //ExStart:SetViewOption
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Load the template document.
            Document doc = new Document(dataDir + "TestFile.doc");
            // Set view option.
            doc.ViewOptions.ViewType = ViewType.PageLayout;
            doc.ViewOptions.ZoomPercent = 50;

            dataDir = dataDir + "TestFile.SetZoom_out_.doc";
            // Save the finished document.
            doc.Save(dataDir);
            //ExEnd:SetViewOption

            Console.WriteLine("\nView option setup successfully.\nFile saved at " + dataDir);
        }
        
    }
}
