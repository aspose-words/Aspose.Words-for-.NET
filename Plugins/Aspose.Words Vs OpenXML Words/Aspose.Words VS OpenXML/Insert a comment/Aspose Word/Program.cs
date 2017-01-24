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
            string File = FilePath + "Insert a comment - Aspose.docx";
            
            // Create an empty document and DocumentBuilder object.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a Comment.
            Comment comment = new Comment(doc);
            // Insert some text into the comment.
            Paragraph commentParagraph = new Paragraph(doc);
            commentParagraph.AppendChild(new Run(doc, "This is comment!!!"));
            comment.AppendChild(commentParagraph);

            // Create CommentRangeStart and CommentRangeEnd.
            int commentId = 0;
            CommentRangeStart start = new CommentRangeStart(doc, commentId);
            CommentRangeEnd end = new CommentRangeEnd(doc, commentId);

            // Insert some text into the document.
            builder.Write("This is text before comment ");
            // Insert comment and comment range start.
            builder.InsertNode(comment);
            builder.InsertNode(start);
            // Insert some more text.
            builder.Write("This is commented text ");
            // Insert end of comment range.
            builder.InsertNode(end);
            // And finaly insert some more text.
            builder.Write("This is text aftr comment");

            // Save output document.
            doc.Save(File);
        }
    }
}
