using Aspose.Words.Saving;
using Aspose.Words.WebExtensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Web_Extensions
{
    class WorkingWithWebExtension
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithWebExtensions();

            UsingWebExtensionTaskPanes(dataDir);
        }

        public static void UsingWebExtensionTaskPanes(string dataDir)
        {
            // ExStart:UsingWebExtensionTaskPanes
            Document doc = new Document();

            TaskPane taskPane = new TaskPane();
            doc.WebExtensionTaskPanes.Add(taskPane);

            taskPane.DockState = TaskPaneDockState.Right;
            taskPane.IsVisible = true;
            taskPane.Width = 300;

            taskPane.WebExtension.Reference.Id = "wa102923726";
            taskPane.WebExtension.Reference.Version = "1.0.0.0";
            taskPane.WebExtension.Reference.StoreType = WebExtensionStoreType.OMEX;
            taskPane.WebExtension.Reference.Store = "th-TH";
            taskPane.WebExtension.Properties.Add(new WebExtensionProperty("mailchimpCampaign", "mailchimpCampaign"));
            taskPane.WebExtension.Bindings.Add(new WebExtensionBinding("UnnamedBinding_0_1506535429545", WebExtensionBindingType.Text, "194740422"));
            
            doc.Save(dataDir + "output.docx", SaveFormat.Docx);
            // ExEnd:UsingWebExtensionTaskPanes 
            Console.WriteLine("\nThe file is saved successfully at " + dataDir);
        }
    }
}
