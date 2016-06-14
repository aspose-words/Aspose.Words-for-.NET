
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Markup;
using System.Drawing;
using Aspose.Words.Drawing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class UpdateContentControls
    {
        public static void Run()
        {            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            SetCurrentStateOfCheckBox(dataDir);   
            //Shows how to modify content controls of type plain text box, drop down list and picture.
            ModifyContentControls(dataDir);
        }
        public static void SetCurrentStateOfCheckBox(string dataDir)
        {
            //ExStart:SetCurrentStateOfCheckBox
            //Open an existing document
            Document doc = new Document(dataDir + "CheckBoxTypeContentControl.docx");

            DocumentBuilder builder = new DocumentBuilder(doc);
            //Get the first content control from the document
            StructuredDocumentTag SdtCheckBox = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

            //StructuredDocumentTag.Checked property gets/sets current state of the Checkbox SDT
            if (SdtCheckBox.SdtType == SdtType.Checkbox)
                SdtCheckBox.Checked = true;

            dataDir = dataDir + "SetCurrentStateOfCheckBox_out_.docx";
            doc.Save(dataDir);
            //ExEnd:SetCurrentStateOfCheckBox
            Console.WriteLine("\nCurrent state fo checkbox setup successfully.\nFile saved at " + dataDir);
        }
        public static void ModifyContentControls(string dataDir)
        {
            //ExStart:ModifyContentControls
            //Open an existing document
            Document doc = new Document(dataDir + "CheckBoxTypeContentControl.docx");

            foreach (StructuredDocumentTag sdt in doc.GetChildNodes(NodeType.StructuredDocumentTag, true))
            {
                if (sdt.SdtType == SdtType.PlainText)
                {
                    sdt.RemoveAllChildren();
                    Paragraph para = sdt.AppendChild(new Paragraph(doc)) as Paragraph;
                    Run run = new Run(doc, "new text goes here");
                    para.AppendChild(run);
                }
                else if (sdt.SdtType == SdtType.DropDownList)
                {
                    SdtListItem secondItem = sdt.ListItems[2];
                    sdt.ListItems.SelectedValue = secondItem;
                }
                else if (sdt.SdtType == SdtType.Picture)
                {
                    Shape shape = (Shape)sdt.GetChild(NodeType.Shape, 0, true);
                    if (shape.HasImage)
                    {
                        shape.ImageData.SetImage(dataDir + "Watermark.png");
                    }
                }
            }


            dataDir = dataDir + "ModifyContentControls_out_.docx";
            doc.Save(dataDir);
            //ExEnd:ModifyContentControls
            Console.WriteLine("\nPlain text box, drop down list and picture content modified successfully.\nFile saved at " + dataDir);
        }
    }
}
