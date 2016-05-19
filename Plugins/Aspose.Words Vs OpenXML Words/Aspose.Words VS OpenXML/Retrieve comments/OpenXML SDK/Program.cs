// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "Retrieve comments.docx";
            
            GetCommentsFromDocument(fileName);
        }
        public static void GetCommentsFromDocument(string fileName)
        {
            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(fileName, false))
            {
                WordprocessingCommentsPart commentsPart =
                    wordDoc.MainDocumentPart.WordprocessingCommentsPart;

                if (commentsPart != null && commentsPart.Comments != null)
                {
                    foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
                    {
                        Console.WriteLine(comment.InnerText);
                    }
                }
            }
        }
    }
}
