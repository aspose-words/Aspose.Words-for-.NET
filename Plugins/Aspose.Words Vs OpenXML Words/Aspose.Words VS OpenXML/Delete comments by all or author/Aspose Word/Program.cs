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
