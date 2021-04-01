using Aspose.Words;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            // The document that the other documents will be appended to.
            Document dstDoc = new Document();

            // All blank documents come with a section with a body with an empty paragraph.
            // Remove all those nodes by using the "RemoveAllChildren" method.
            dstDoc.RemoveAllChildren();

            Document doc1 = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc1);

            builder.Writeln("Hello world! This is the first document.");

            Document doc2 = new Document();
            builder = new DocumentBuilder(doc2);

            builder.Writeln("Hello again! This is the second document.");

            dstDoc.AppendDocument(doc1, ImportFormatMode.UseDestinationStyles);
            dstDoc.AppendDocument(doc2, ImportFormatMode.UseDestinationStyles);

            // Each appended document starts a new section.
            // Make sure that none of the sections link to headers that came from other documents.
            for (int i = 0; i < dstDoc.Sections.Count; i++)
                if (i > 0)
                    dstDoc.Sections[i].HeadersFooters.LinkToPrevious(false);

            dstDoc.Save("Joining Documents Together.docx");
        }
    }
}
