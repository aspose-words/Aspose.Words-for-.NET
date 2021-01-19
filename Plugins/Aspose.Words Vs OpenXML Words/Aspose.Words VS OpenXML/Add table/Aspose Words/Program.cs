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
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 2x2 table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1 Content.");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content.");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Row 2, Cell 1 Content");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content.");
            builder.EndRow();
            builder.EndTable();

            // Save the document to the local file system.
            string FilePath = @"..\..\..\..\Sample Files\";
            doc.Save(FilePath + "Add Table - Aspose.docx");
        }
    }
}
