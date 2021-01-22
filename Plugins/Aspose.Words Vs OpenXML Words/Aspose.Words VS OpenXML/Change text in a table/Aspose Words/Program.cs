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
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

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
            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = true, 
                FindWholeWordsOnly = true
            };

            table.Rows[1].Cells[2].Range.Replace("Mr", "test", options);

            doc.Save(File);
        }
    }
}
