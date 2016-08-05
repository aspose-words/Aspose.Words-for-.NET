using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Fonts;
using Aspose.Words;
using Aspose.Words.Saving;
namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class ReceiveNotificationsOfFont 
    {
        public static void Run()
        {
            //ExStart:ReceiveNotificationsOfFonts 
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            Document doc = new Document(dataDir + "Rendering.doc");

            FontSettings FontSettings = new FontSettings();          

            // We can choose the default font to use in the case of any missing fonts.
            FontSettings.DefaultFontName = "Arial";
            // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
            // find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
            // font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
            FontSettings.SetFontsFolder(string.Empty, false);

            // Create a new class implementing IWarningCallback which collect any warnings produced during document save.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();

            doc.WarningCallback = callback;
            // Set font settings
            doc.FontSettings = FontSettings;
            string path = dataDir + "Rendering.MissingFontNotification_out_.pdf";
            // Pass the save options along with the save path to the save method.
            doc.Save(path);
            //ExEnd:ReceiveNotificationsOfFonts 
            Console.WriteLine("\nReceive notifications of font substitutions by using IWarningCallback processed.\nFile saved at " + path);

            ReceiveWarningNotification(doc, dataDir);
        }
        private static void ReceiveWarningNotification(Document doc, string dataDir)
        {
            //ExStart:ReceiveWarningNotification 
            // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occured during rendering
            // are stored until the document save and then sent to the appropriate WarningCallback.
            doc.UpdatePageLayout();

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();

            doc.WarningCallback = callback;
            dataDir = dataDir + "Rendering.FontsNotificationUpdatePageLayout_out_.pdf";
            // Even though the document was rendered previously, any save warnings are notified to the user during document save.
            doc.Save(dataDir);
            //ExEnd:ReceiveWarningNotification  
        }
        //ExStart:HandleDocumentWarnings
        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document procssing. The callback can be set to listen for warnings generated during document
            /// load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // We are only interested in fonts being substituted.
                if (info.WarningType == WarningType.FontSubstitution)
                {
                    Console.WriteLine("Font substitution: " + info.Description);
                }
            }
        }
        //ExEnd:HandleDocumentWarnings
    }
}
