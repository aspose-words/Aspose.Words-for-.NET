using System;
using System.IO;
using Aspose.Words;

namespace _01._04_TrackChangesInDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for an Aspose.Words license file in the local file system and apply it, if it exists.
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                Aspose.Words.License license = new Aspose.Words.License();

                // Use the license from the bin/debug/ Folder.
                license.SetLicense("Aspose.Words.lic");
            }
            
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Normal editing of the document does not count as a revision.
            builder.Write("This does not count as a revision. ");
            
            // To register our edits as revisions, we need to declare an author, and then start tracking them.
            doc.StartTrackRevisions("John Doe", DateTime.Now);

            builder.Write("This is an insert revision. ");

            // Accept the revision to assimilate its contents into the document's body.
            doc.Revisions[0].Accept();

            builder.Write("This is another insert revision. ");

            // Reject an insert revision to leave it out of the document's body and discard its contents.
            doc.Revisions[0].Reject();

            // Stop tracking revisions to continue editing the document as normal.
            doc.StopTrackRevisions(); 

            builder.Write("This does not count as a revision.");

            doc.Save("TrackChangesInDocuments.docx");
        }
    }
}
