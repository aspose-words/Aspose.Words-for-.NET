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
            string filepath = @"..\..\..\..\Sample Files\";
            string fullFilename = filepath + "Create and add a paragraph style - Aspose.docx";
            
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            Style style = doc.Styles.Add(StyleType.Paragraph, "MyStyle");
            Aspose.Words.Font font = builder.Font;
            font.Bold = true;
            font.Color = System.Drawing.Color.Blue;
            font.Italic = true;
            font.Name = "Arial";
            font.Size = 24;
            font.Spacing = 5;
            font.Underline = Underline.Double;

            builder.ParagraphFormat.Style = doc.Styles["MyStyle"];

            builder.MoveToDocumentEnd();
            builder.Writeln("This string is formatted using the new style.");

            doc.Save(fullFilename);
        }
    }
}
