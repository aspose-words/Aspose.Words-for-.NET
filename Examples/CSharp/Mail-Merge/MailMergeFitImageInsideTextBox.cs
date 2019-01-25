using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MailMergeFitImageInsideTextBox
    {
        public static void Run()
        {
            // ExStart:MailMergeFitImageInsideTextBox
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting();
            Document doc = new Document(dataDir + "MailMerge.TextBox.FitImage.doc");


            // Set up the event handler for image fields.
            doc.MailMerge.FieldMergingCallback = new HandleFieldMergingCallback();
            doc.MailMerge.UseNonMergeFields = true;

            doc.MailMerge.Execute(
               new string[] { "FullName", "Company", "Address", "Address2", "City", "image__c" },
               new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London", "https://blog.aspose.com/wp-content/uploads/sites/2/2011/08/Aspose.Words_.Express-Settings.jpg" });

            dataDir = dataDir + "MailMerge.TextBox.FitImage_out.doc";
            doc.Save(dataDir);
            // ExEnd:MailMergeFitImageInsideTextBox
            Console.WriteLine("\nMail merge fit image inside textbox performed successfully.\nFile saved at " + dataDir);
        }

        // ExStart:HandleFieldMergingCallback 
        public class HandleFieldMergingCallback : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                // Do nothing.
            }

            /// <summary>
            /// This is called when mail merge engine encounters Image:XXX merge field in the document.
            /// </summary>
            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
            {
                if (e.Field.GetFieldCode().Contains("Image:"))
                {
                    Shape textBox = (Shape)e.Field.Start.GetAncestor(NodeType.Shape);
                    if (textBox.ShapeType == ShapeType.TextBox)
                    {
                        textBox.TextBox.InternalMarginLeft = 0;
                        textBox.TextBox.InternalMarginRight = 0;
                        textBox.TextBox.InternalMarginTop = 0;
                        textBox.TextBox.InternalMarginBottom = 0;

                        e.ImageWidth = new MergeFieldImageDimension(textBox.Width);
                        e.ImageHeight = new MergeFieldImageDimension(textBox.Height);
                    }
                }
            }
        }
        // ExEnd:HandleFieldMergingCallback
    }
}
