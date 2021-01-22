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
            string File = FilePath + "Insert a comment - Aspose.docx";
            
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Comment comment = new Comment(doc);

            // Insert some text into the comment.
            Paragraph commentParagraph = new Paragraph(doc);
            commentParagraph.AppendChild(new Run(doc, "This is comment!!!"));
            comment.AppendChild(commentParagraph);

            // Create a "CommentRangeStart" and "CommentRangeEnd".
            int commentId = 0;
            CommentRangeStart start = new CommentRangeStart(doc, commentId);
            CommentRangeEnd end = new CommentRangeEnd(doc, commentId);

            builder.Write("This text is before the comment. ");

            // Insert comment and comment range start.
            builder.InsertNode(comment);
            builder.InsertNode(start);

            // Insert some more text.
            builder.Write("This text is commented. ");

            // Insert end of comment range.
            builder.InsertNode(end);

            builder.Write("This text is after the comment.");

            doc.Save(File);
        }
    }
}
