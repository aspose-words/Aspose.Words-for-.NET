using System;
using Aspose.Words;
using Aspose.Words.WebExtensions;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Document
{
    internal class WorkingWithWebExtension : DocsExamplesBase
    {
        [Test]
        public void UsingWebExtensionTaskPanes()
        {
            //ExStart:UsingWebExtensionTaskPanes
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
            taskPane.WebExtension.Bindings.Add(new WebExtensionBinding("UnnamedBinding_0_1506535429545",
                WebExtensionBindingType.Text, "194740422"));

            doc.Save(ArtifactsDir + "WorkingWithWebExtension.UsingWebExtensionTaskPanes.docx");
            //ExEnd:UsingWebExtensionTaskPanes
            
            //ExStart:GetListOfAddins
            doc = new Document(ArtifactsDir + "WorkingWithWebExtension.UsingWebExtensionTaskPanes.docx");
            
            Console.WriteLine("Task panes sources:\n");

            foreach (TaskPane taskPaneInfo in doc.WebExtensionTaskPanes)
            {
                WebExtensionReference reference = taskPaneInfo.WebExtension.Reference;
                Console.WriteLine($"Provider: \"{reference.Store}\", version: \"{reference.Version}\", catalog identifier: \"{reference.Id}\";");
            }
            //ExEnd:GetListOfAddins
        }
    }
}