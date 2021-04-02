// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class JoiningMultipleDocuments : TestUtil
    {
        [Test]
        public static void JoiningMultipleDocumentsFeature()
        {
            // The document that the other documents will be appended to.
            Document dstDoc = new Document();

            // We should call this method to clear this document of any existing content.
            dstDoc.RemoveAllChildren();

            int recordCount = 1;
            for (int i = 1; i <= recordCount; i++)
            {
                // Open the document to join.
                Document srcDoc = new Document(MyDir + "Joining mutiple documents 1.docx");
                // Append the source document at the end of the destination document.
                dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

                Document doc2 = new Document(MyDir + "Joining mutiple documents 2.docx");
                dstDoc.AppendDocument(doc2, ImportFormatMode.UseDestinationStyles);
                
                // If this is the second document or above being appended then unlink all headers footers in this section
                // from the headers and footers of the previous section.
                if (i > 1)
                    dstDoc.Sections[i].HeadersFooters.LinkToPrevious(false);
            }

            dstDoc.Save(ArtifactsDir + "Joining multiple documents - Aspose.Words.docx");
        }
    }
}
