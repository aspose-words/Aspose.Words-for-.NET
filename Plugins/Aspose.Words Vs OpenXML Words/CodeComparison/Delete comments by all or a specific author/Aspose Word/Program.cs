// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;
namespace Aspose_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc=new Document("Delete comments by all or by an specific author.docx");
            RemoveComments(doc, "");
        }
        public static void RemoveComments(Document doc, string authorName)
        {
            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            if (authorName == "")
            {
                // Remove all comments.
                comments.Clear();
            }
            else
            {
                // Look through all comments and remove those written by the authorName author.
                for (int i = comments.Count - 1; i >= 0; i--)
                {
                    Comment comment = (Comment)comments[i];
                    if (comment.Author == authorName)
                        comment.Remove();
                }
            }
            doc.Save(@"F:\Dropbox\Personal\Aspose Vs OpenXML\Files\Delete comments by all or by an specific author.docx");
        }
    }
}
