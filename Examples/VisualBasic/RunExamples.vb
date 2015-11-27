Imports System.IO
Imports DocumentExplorerExample
Imports LINQ

Module RunExamples

    Sub Main()
        Console.WriteLine("Open RunExamples.vb. In Main() method, Un-comment the example that you want to run")
        Console.WriteLine("=====================================================")


        '' =====================================================
        '' =====================================================
        '' Quick Start
        '' =====================================================
        '' =====================================================
        AppendDocuments.Run()
        'ApplyLicense.Run()
        'Doc2Pdf.Run()
        'FindAndReplace.Run()
        'HelloWorld.Run()
        'LoadAndSaveToDisk.Run()
        'LoadAndSaveToStream.Run()
        'SimpleMailMerge.Run()
        'UpdateFields.Run()
        'WorkingWithNodes.Run()

        '' =====================================================
        '' =====================================================
        '' Loading and Saving
        '' =====================================================
        '' =====================================================

        'CheckFormat.Run()
        'SplitIntoHtmlPages.Run()
        'LoadTxt.Run()
        'PageSplitter.Run()
        'ImageToPdf.Run()

        '' =====================================================
        '' =====================================================
        '' Programming with Documents
        '' =====================================================
        '' =====================================================

        '' Joining and Appending
        '' =====================================================
        'SimpleAppendDocument.Run()
        'KeepSourceFormatting.Run()
        'UseDestinationStyles.Run()
        'JoinContinuous.Run()
        'JoinNewPage.Run()
        'RestartPageNumbering.Run()
        'LinkHeadersFooters.Run()
        'UnlinkHeadersFooters.Run()
        'RemoveSourceHeadersFooters.Run()
        'DifferentPageSetup.Run()
        'ConvertNumPageFields.Run()
        'ListUseDestinationStyles.Run()
        'ListKeepSourceFormatting.Run()
        'KeepSourceTogether.Run()
        'BaseDocument.Run()
        'UpdatePageLayout.Run()
        'AppendDocumentManually.Run()
        'PrependDocument.Run()

        '' Find and Replace
        '' =====================================================
        'FindAndHighlight.Run()
        'ReplaceTextWithField.Run()

        '' Bookmarks
        '' =====================================================
        'CopyBookmarkedText.Run()
        'UntangleRowBookmarks.Run()

        '' Comments
        '' =====================================================
        'ProcessComments.Run()

        '' Document
        '' =====================================================
        'ExtractContent.Run()
        'PageNumbersOfNodes.Run()
        'RemoveBreaks.Run()

        '' Fields
        '' =====================================================
        'InsertNestedFields.Run()
        'RemoveField.Run()
        'ConvertFieldsInDocument.Run()
        'ConvertFieldsInBody.Run()
        'ConvertFieldsInParagraph.Run()

        '' Images
        '' =====================================================
        'AddImageToEachPage.Run()
        'AddWatermark.Run()
        'CompressImages.Run()

        '' Styles
        '' =====================================================
        'ExtractContentBasedOnStyles.Run()

        '' Tables
        '' =====================================================
        'AutoFitTableToWindow.Run()
        'AutoFitTableToContents.Run()
        'AutoFitTableToFixedColumnWidths.Run()

        '' =====================================================
        '' =====================================================
        '' MailMerge and Reporting
        '' =====================================================
        '' =====================================================

        'ApplyCustomLogicToEmptyRegions.Run()
        'LINQtoXMLMailMerge.Run()
        'MailMergeFormFields.Run()
        'MultipleDocsInMailMerge.Run()
        'NestedMailMerge.Run()
        'RemoveEmptyRegions.Run()
        'XMLMailMerge.Run()

        '' =====================================================
        '' =====================================================
        '' Rendering and Printing
        '' =====================================================
        '' =====================================================

        'DocumentLayoutHelper.Run()
        'EnumerateLayoutElements.Run()
        'DocumentPreviewAndPrint.Run()
        'ImageColorFilters.Run()
        'RenderShape.Run()
        'SaveAsMultipageTiff.Run()

        '' =====================================================
        '' =====================================================
        '' Viewers and Visualizers
        '' =====================================================
        '' =====================================================

        'MainForm.Run()

        '' =====================================================
        '' =====================================================
        '' LINQ
        '' =====================================================
        '' =====================================================
        'LINQ.HelloWorld.Run()
        'SingleRow.Run()
        'InParagraphList.Run()
        'BulletedList.Run()
        'NumberedList.Run()
        'MulticoloredNumberedList.Run()
        'CommonList.Run()
        'InTableList.Run()
        'InTableAlternateContent.Run()
        'CommonMasterDetail.Run()
        'InTableMasterDetail.Run()
        'InTableWithFilteringGroupingSorting.Run()
        'PieChart.Run()
        'ScatterChart.Run()
        'BubbleChart.Run()
        'ChartWithFilteringGroupingOrdering.Run()

        ' Stop before exiting
        Console.WriteLine(vbNewLine + vbNewLine + "Program Finished. Press any key to exit....")
        Console.ReadKey()
    End Sub

    Public Function GetDataDir_LINQ() As [String]
        Return Path.GetFullPath("../../LINQ/Data/")
    End Function

    Public Function GetDataDir_LoadingAndSaving() As [String]
        Return Path.GetFullPath("../../Loading-and-Saving/Data/")
    End Function

    Public Function GetDataDir_JoiningAndAppending() As [String]
        Return Path.GetFullPath("../../Programming-Documents/Joining-Appending/Data/")
    End Function

    Public Function GetDataDir_FindAndReplace() As [String]
        Return Path.GetFullPath("../../Programming-Documents/Find-Replace/Data/")
    End Function

    Public Function GetDataDir_WorkingWithBookmarks() As [String]
        Return Path.GetFullPath("../../Programming-Documents/Bookmarks/Data/")
    End Function

    Public Function GetDataDir_WorkingWithComments() As [String]
        Return Path.GetFullPath("../../Programming-Documents/Comments/Data/")
    End Function

    Public Function GetDataDir_WorkingWithDocument() As [String]
        Return Path.GetFullPath("../../Programming-Documents/Document/Data/")
    End Function

    Public Function GetDataDir_WorkingWithFields() As [String]
        Return Path.GetFullPath("../../Programming-Documents/Fields/Data/")
    End Function

    Public Function GetDataDir_WorkingWithImages() As [String]
        Return Path.GetFullPath("../../Programming-Documents/Images/Data/")
    End Function

    Public Function GetDataDir_WorkingWithStyles() As [String]
        Return Path.GetFullPath("../../Programming-Documents/Styles/Data/")
    End Function

    Public Function GetDataDir_WorkingWithTables() As [String]
        Return Path.GetFullPath("../../Programming-Documents/Tables/Data/")
    End Function

    Public Function GetDataDir_MailMergeAndReporting() As [String]
        Return Path.GetFullPath("../../Mail-Merge/Data/")
    End Function

    Public Function GetDataDir_QuickStart() As [String]
        Return Path.GetFullPath("../../Quick-Start/Data/")
    End Function

    Public Function GetDataDir_RenderingAndPrinting() As [String]
        Return Path.GetFullPath("../../Rendering-Printing/Data/")
    End Function

    Public Function GetDataDir_ViewersAndVisualizers() As [String]
        Return Path.GetFullPath("../../Viewers-Visualizers/Data/")
    End Function

End Module
