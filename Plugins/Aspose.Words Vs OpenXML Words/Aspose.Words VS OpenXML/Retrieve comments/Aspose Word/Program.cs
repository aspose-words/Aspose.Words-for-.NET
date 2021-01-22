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
using System;
using System.Collections;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "Retrieve comments.docx";
            
            Document doc = new Document(fileName);
            ExtractComments(doc);
        }

        public static void ExtractComments(Document doc)
        {
            ArrayList collectedComments = new ArrayList();
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            // Look through all comments and gather information about them.
            foreach (Comment comment in comments)
                collectedComments.Add(comment.Author + " " + comment.DateTime + " " + comment.ToString(SaveFormat.Text));

            foreach (string collectedComment in collectedComments)
                Console.WriteLine(collectedComment);
        }
    }
}
