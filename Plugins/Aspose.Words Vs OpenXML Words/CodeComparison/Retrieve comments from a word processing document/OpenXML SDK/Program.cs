// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXML_SDK
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "Retrieve comments from word processing document.docx";
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
