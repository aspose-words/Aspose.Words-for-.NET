using CSharp.Loading_Saving;
using CSharp.Mail_Merge;
using CSharp.Programming_Documents.Find_and_Replace;
using CSharp.Programming_Documents.Joining_and_Appending;
using CSharp.Programming_Documents.Bookmarks;
using CSharp.Programming_Documents.Comments;
using CSharp.Programming_Documents.Working_With_Document;
using CSharp.Programming_Documents.Working_with_Fields;
using CSharp.Programming_Documents.Working_with_Images;
using CSharp.Programming_Documents.Working_with_Styles;
using CSharp.Programming_Documents.Working_with_Tables;
using CSharp.Quick_Start;
using CSharp.Rendering_and_Printing;
using CSharp.LINQ;
using DocumentExplorerExample;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CSharp
{
    class RunExamples
    {
        [STAThread]
        public static void Main()
        {
            Console.WriteLine("Open RunExamples.cs. In Main() method, Un-comment the example that you want to run");
            Console.WriteLine("=====================================================");
            // Un-comment the one you want to try out

            // =====================================================
            // =====================================================
            // Quick Start
            // =====================================================
            // =====================================================

            //AppendDocuments.Run();
            //ApplyLicense.Run();
            //Doc2Pdf.Run();
            //FindAndReplace.Run();
            //HelloWorld.Run();
            //LoadAndSaveToDisk.Run();
            //LoadAndSaveToStream.Run();
            //SimpleMailMerge.Run();
            //UpdateFields.Run();
            //WorkingWithNodes.Run();

            //// =====================================================
            //// =====================================================
            //// Loading and Saving
            //// =====================================================
            //// =====================================================

            //CheckFormat.Run();
            //SplitIntoHtmlPages.Run();
            //LoadTxt.Run();
            //PageSplitter.Run();
            //ImageToPdf.Run();

            //// =====================================================
            //// =====================================================
            //// Programming with Documents
            //// =====================================================
            //// =====================================================

            //// Joining and Appending
            //// =====================================================
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

            //// Find and Replace
            //// =====================================================
            //FindAndHighlight.Run();
            //ReplaceTextWithField.Run();

            //// Bookmarks
            //// =====================================================
            //CopyBookmarkedText.Run();
            //UntangleRowBookmarks.Run();

            //// Comments
            //// =====================================================
            //ProcessComments.Run();

            //// Document
            //// =====================================================
            //ExtractContentBetweenParagraphs.Run();
            //ExtractContentBetweenBlockLevelNodes.Run();
            //ExtractContentBetweenParagraphStyles.Run();
            //ExtractContentBetweenRuns.Run();
            //ExtractContentUsingField.Run();
            //ExtractContentBetweenBookmark.Run();
            //ExtractContentBetweenCommentRange.Run();
            //PageNumbersOfNodes.Run();
            //RemoveBreaks.Run();

            //// Fields
            //// =====================================================
            //InsertNestedFields.Run();
            //RemoveField.Run();
            //ConvertFieldsInDocument.Run();
            //ConvertFieldsInBody.Run();
            //ConvertFieldsInParagraph.Run();

            //// Images
            //// =====================================================
            //AddImageToEachPage.Run();
            //AddWatermark.Run();
            //CompressImages.Run();

            //// Styles
            //// =====================================================
            //ExtractContentBasedOnStyles.Run();

            //// Tables
            //// =====================================================
            //AutoFitTableToWindow.Run();
            //AutoFitTableToContents.Run();
            //AutoFitTableToFixedColumnWidths.Run();

            //// =====================================================
            //// =====================================================
            //// MailMerge and Reporting
            //// =====================================================
            //// =====================================================

            //ApplyCustomLogicToEmptyRegions.Run();
            //LINQtoXMLMailMerge.Run();
            //MailMergeFormFields.Run();
            //MultipleDocsInMailMerge.Run();
            //NestedMailMerge.Run();
            //RemoveEmptyRegions.Run();
            //XMLMailMerge.Run();            

            //// =====================================================
            //// =====================================================
            //// Rendering and Printing
            //// =====================================================
            //// =====================================================

            //DocumentLayoutHelper.Run();
            //EnumerateLayoutElements.Run();
            //DocumentPreviewAndPrint.Run();
            //ImageColorFilters.Run();
            //RenderShape.Run();
            //SaveAsMultipageTiff.Run();
            ReadActiveXControlProperties.Run();

            //// =====================================================
            //// =====================================================
            //// Viewers and Visualizers
            //// =====================================================
            //// =====================================================

            //MainForm.Run();

            //// =====================================================
            //// =====================================================
            //// LINQ
            //// =====================================================
            //// =====================================================

            //CSharp.LINQ.HelloWorld.Run();
            //SingleRow.Run();
            //InParagraphList.Run();
            //BulletedList.Run();
            //NumberedList.Run();
            //MulticoloredNumberedList.Run();
            //CommonList.Run();
            //InTableList.Run();
            //InTableAlternateContent.Run();
            //CommonMasterDetail.Run();
            //InTableMasterDetail.Run();
            //InTableWithFilteringGroupingSorting.Run();
            //PieChart.Run();
            //ScatterChart.Run();
            //BubbleChart.Run();
            //ChartWithFilteringGroupingOrdering.Run();
         
            // Stop before exiting
            Console.WriteLine("\n\nProgram Finished. Press any key to exit....");
            Console.ReadKey();
        }

        public static String GetDataDir_LINQ()
        {
            return Path.GetFullPath("../../LINQ/Data/");
        }
        public static String GetDataDir_LoadingAndSaving()
        {
            return Path.GetFullPath("../../Loading-and-Saving/Data/");
        }

        public static String GetDataDir_JoiningAndAppending()
        {
            return Path.GetFullPath("../../Programming-Documents/Joining-Appending/Data/");
        }

        public static String GetDataDir_FindAndReplace()
        {
            return Path.GetFullPath("../../Programming-Documents/Find-Replace/Data/");
        }

        public static String GetDataDir_WorkingWithBookmarks()
        {
            return Path.GetFullPath("../../Programming-Documents/Bookmarks/Data/");
        }

        public static String GetDataDir_WorkingWithComments()
        {
            return Path.GetFullPath("../../Programming-Documents/Comments/Data/");
        }

        public static String GetDataDir_WorkingWithDocument()
        {
            return Path.GetFullPath("../../Programming-Documents/Document/Data/");
        }

        public static String GetDataDir_WorkingWithFields()
        {
            return Path.GetFullPath("../../Programming-Documents/Fields/Data/");
        }

        public static String GetDataDir_WorkingWithImages()
        {
            return Path.GetFullPath("../../Programming-Documents/Images/Data/");
        }

        public static String GetDataDir_WorkingWithStyles()
        {
            return Path.GetFullPath("../../Programming-Documents/Styles/Data/");
        }

        public static String GetDataDir_WorkingWithTables()
        {
            return Path.GetFullPath("../../Programming-Documents/Tables/Data/");
        }

        public static String GetDataDir_MailMergeAndReporting()
        {
            return Path.GetFullPath("../../Mail-Merge/Data/");
        }

        public static String GetDataDir_QuickStart()
        {
            return Path.GetFullPath("../../Quick-Start/Data/");
        }

        public static String GetDataDir_RenderingAndPrinting()
        {
            return Path.GetFullPath("../../Rendering-Printing/Data/");
        }

        public static String GetDataDir_ViewersAndVisualizers()
        {
            return Path.GetFullPath("../../Viewers-Visualizers/Data/");
        }
    }
}
