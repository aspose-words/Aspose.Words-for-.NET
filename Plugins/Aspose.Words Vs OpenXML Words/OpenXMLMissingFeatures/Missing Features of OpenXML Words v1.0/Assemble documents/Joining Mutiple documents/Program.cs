// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JoiningMutipleDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            // The document that the other documents will be appended to.
            Document dstDoc = new Document();

            // We should call this method to clear this document of any existing content.
            dstDoc.RemoveAllChildren();

            int recordCount = 1;
            for (int i = 1; i <= recordCount; i++)
            {
                // Open the document to join.
                Document srcDoc = new Document(MyDir + "src.doc");

                // Append the source document at the end of the destination document.
                dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
                Document doc2 = new Document(MyDir + "srcDocument.doc");
                dstDoc.AppendDocument(doc2, ImportFormatMode.UseDestinationStyles);
                // If this is the second document or above being appended then unlink all headers footers in this section
                // from the headers and footers of the previous section.
                if (i > 1)
                    dstDoc.Sections[i].HeadersFooters.LinkToPrevious(false);
            }
            dstDoc.Save(MyDir + "MutipleJoinedDocument.doc");
        }
    }
}
