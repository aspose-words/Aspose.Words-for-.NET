// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string SrcFileName = FilePath + "Joining Mutiple documents 1.docx";
            string DestFileName = FilePath + "Joining Mutiple documents 2.docx";
            
            // The document that the other documents will be appended to.
            Document dstDoc = new Document();

            // We should call this method to clear this document of any existing content.
            dstDoc.RemoveAllChildren();

            int recordCount = 1;
            for (int i = 1; i <= recordCount; i++)
            {
                // Open the document to join.
                Document srcDoc = new Document(SrcFileName);

                // Append the source document at the end of the destination document.
                dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
                Document doc2 = new Document(DestFileName);
                dstDoc.AppendDocument(doc2, ImportFormatMode.UseDestinationStyles);
                // If this is the second document or above being appended then unlink all headers footers in this section
                // from the headers and footers of the previous section.
                if (i > 1)
                    dstDoc.Sections[i].HeadersFooters.LinkToPrevious(false);
            }
            dstDoc.Save(DestFileName);
        }
    }
}
