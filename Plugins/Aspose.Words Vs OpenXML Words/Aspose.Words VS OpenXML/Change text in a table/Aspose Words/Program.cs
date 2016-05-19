// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;
using Aspose.Words.Tables;
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
            string File = FilePath + "Change text in a table - Aspose.docx";
            
            Document doc = new Document(File);

            // Get the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
             
            // Replace any instances of our string in the last cell of the table only.
            table.Rows[1].Cells[2].Range.Replace("Mr", "test", true, true);
            doc.Save(File);
        }
    }
}
