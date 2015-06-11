using CSharp.Programming_With_Documents.Find_and_Replace;
using CSharp.Programming_With_Documents.Joining_and_Appending;
using CSharp.Programming_With_Documents.Working_with_Bookmarks;
using CSharp.Programming_With_Documents.Working_with_Comments;
using CSharp.Programming_With_Documents.Working_with_Document;
using CSharp.Programming_With_Documents.Working_with_Fields;
using CSharp.Programming_With_Documents.Working_with_Images;
using CSharp.Programming_With_Documents.Working_with_Styles;
using CSharp.Programming_With_Documents.Working_with_Tables;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CSharp.Programming_With_Documents
{
    class _RunExamples
    {
        public static void Main()
        {
            // Run the examples. Un-comment the one you want to run

            // Joining and Appending
            // =====================================================
            //SimpleAppendDocument.Run();
            //KeepSourceFormatting.Run();
            //UseDestinationStyles.Run();
            //JoinContinuous.Run();
            //JoinNewPage.Run();
            //RestartPageNumbering.Run();
            //LinkHeadersFooters.Run();
            //UnlinkHeadersFooters.Run();
            //RemoveSourceHeadersFooters.Run();
            //DifferentPageSetup.Run();
            //ConvertNumPageFields.Run();
            //ListUseDestinationStyles.Run();
            //ListKeepSourceFormatting.Run();
            //KeepSourceTogether.Run();
            //BaseDocument.Run();
            //UpdatePageLayout.Run();
            //AppendDocumentManually.Run();
            //PrependDocument.Run();

            // Find and Replace
            // =====================================================
            //FindAndHighlight.Run();
            //ReplaceTextWithField.Run();

            // Working with Bookmarks
            // =====================================================
            //CopyBookmarkedText.Run();
            //UntangleRowBookmarks.Run();

            // Working with Bookmarks
            // =====================================================
            //ProcessComments.Run();

            // Working with Document
            // =====================================================
            //ExtractContentBetweenParagraphs.Run();
            //ExtractContentBetweenBlockLevelNodes.Run();
            //ExtractContentBetweenParagraphStyles.Run();
            //ExtractContentBetweenRuns.Run();
            //ExtractContentUsingField.Run();
            //ExtractContentBetweenBookmark.Run();
            //ExtractContentBetweenCommentRange.Run();
            //PageNumbersOfNodes.Run();
            //RemoveBreaks.Run();

            // Working with Fields
            // =====================================================
            //InsertNestedFields.Run();
            //RemoveField.Run();
            //ConvertFieldsInDocument.Run();
            //ConvertFieldsInBody.Run();
            //ConvertFieldsInParagraph.Run();

            // Working with Images
            // =====================================================
            //AddImageToEachPage.Run();
            //AddWatermark.Run();
            //CompressImages.Run();

            // Working with Styles
            // =====================================================
            //ExtractContentBasedOnStyles.Run();

            // Working with Styles
            // =====================================================
            AutoFitTableToWindow.Run();
            AutoFitTableToContents.Run();
            AutoFitTableToFixedColumnWidths.Run();
            
            // Stop before exiting
            Console.WriteLine("\n\nProgram Finished. Press any key to exit....");
            Console.ReadKey();
        }

        public static String GetDataDir_JoiningAndAppending()
        {
            return Path.GetFullPath("../../Programming-With-Documents/Joining-and-Appending/Data/");
        }

        public static String GetDataDir_FindAndReplace()
        {
            return Path.GetFullPath("../../Programming-With-Documents/Find-and-Replace/Data/");
        }

        public static String GetDataDir_WorkingWithBookmarks()
        {
            return Path.GetFullPath("../../Programming-With-Documents/Working-with-Bookmarks/Data/");
        }

        public static String GetDataDir_WorkingWithComments()
        {
            return Path.GetFullPath("../../Programming-With-Documents/Working-with-Comments/Data/");
        }

        public static String GetDataDir_WorkingWithDocument()
        {
            return Path.GetFullPath("../../Programming-With-Documents/Working-with-Document/Data/");
        }

        public static String GetDataDir_WorkingWithFields()
        {
            return Path.GetFullPath("../../Programming-With-Documents/Working-with-Fields/Data/");
        }

        public static String GetDataDir_WorkingWithImages()
        {
            return Path.GetFullPath("../../Programming-With-Documents/Working-with-Images/Data/");
        }

        public static String GetDataDir_WorkingWithStyles()
        {
            return Path.GetFullPath("../../Programming-With-Documents/Working-with-Styles/Data/");
        }

        public static String GetDataDir_WorkingWithTables()
        {
            return Path.GetFullPath("../../Programming-With-Documents/Working-with-Tables/Data/");
        }
    }
}
