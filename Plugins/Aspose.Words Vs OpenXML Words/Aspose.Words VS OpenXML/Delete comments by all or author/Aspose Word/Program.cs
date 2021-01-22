// Copyright (c) Aspose 2002-2021. All Rights Reserved.

/*
    This project uses NuGet's Automatic Package Restore feature to 
    resolve the Aspose.Words for .NET API reference when the project is built. 
    Please visit https://docs.nuget.org/consume/nuget-faq for more information. 

    If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API 
    from http://www.aspose.com/downloads, install it, and then add a reference to it to this project. 

    For any issues, questions or suggestions, please visit the Aspose Forums: https://forum.aspose.com/
*/

using Aspose.Words;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string File = FilePath + "Delete comments - Aspose.docx";

            RemoveComments(File, "");
        }
        public static void RemoveComments(string File, string authorName)
        {
            Document doc = new Document(File);

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
            doc.Save(File);
        }
    }
}
