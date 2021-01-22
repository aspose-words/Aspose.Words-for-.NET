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
