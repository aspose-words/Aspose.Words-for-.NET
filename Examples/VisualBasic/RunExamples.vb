Imports System.IO
Imports Aspose.Words.Examples.VisualBasic.DocumentExplorerExample
Imports Aspose.Words.Examples.VisualBasic.LINQ

Module RunExamples

    Sub Main()
        Console.WriteLine("Open RunExamples.vb. " & vbLf & "In Main() method uncomment the example that you want to run.")
        Console.WriteLine("=====================================================")

        ' Uncomment the one you want to try out

        '' =====================================================
        '' =====================================================
        '' Quick Start
        '' =====================================================
        '' =====================================================
        'AppendDocuments.Run()
        'ApplyLicense.Run()
        'FindAndReplace.Run()
        'HelloWorld.Run()
        'UpdateFields.Run()
        'WorkingWithNodes.Run()

        '' =====================================================
        '' =====================================================
        '' Loading and Saving
        '' =====================================================
        '' =====================================================

        'OpenEncryptedDocument.Run()
        'LoadAndSaveToDisk.Run()
        'LoadAndSaveToStream.Run()
        'CreateDocument.Run()
        'CheckFormat.Run()
        'SplitIntoHtmlPages.Run()
        'LoadTxt.Run()
        'PageSplitter.Run()
        'ImageToPdf.Run()
        'SpecifySaveOption.Run()
        'AccessAndVerifySignature.Run()
        'Doc2Pdf.Run()
        'DigitallySignedPdf.Run()
        'ConvertDocumentToByte.Run()
        'ConvertDocumentToEPUB.Run()
        'ConvertDocumentToHtmlWithRoundtrip.Run()
        'DetectDocumentSignatures.Run()
        'SaveAsTxt.Run()

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
        'ReplaceWithString.Run()
        'ReplaceWithRegex.Run()
        'ReplaceWithEvaluator.Run()

        '' Bookmarks
        '' =====================================================
        'CopyBookmarkedText.Run()
        'UntangleRowBookmarks.Run()
        'BookmarkTable.Run()
        'BookmarkNameAndText.Run()
        'AccessBookmarks.Run()
        'CreateBookmark.Run()

        '' Comments
        '' =====================================================
        'ProcessComments.Run()
        'AddComments.Run()
        'AnchorComment.Run()

        '' ConvertUtil
        '' =====================================================
        'UtilityClasses.Run()

        '' Theme
        '' =====================================================
        'ManipulateThemeProperties.Run()

        '' Node
        '' =====================================================
        'ExNode.Run()

        '' Charts
        '' =====================================================
        'CreateColumnChart.Run()
        'InsertScatterChart.Run()
        'InsertAreaChart.Run()
        'InsertBubbleChart.Run()
        'CreateChartUsingShape.Run()
        'WorkWithChartDataLabel.Run()
        'WorkWithSingleChartDataPoint.Run()
        'WorkWithSingleChartSeries.Run()

        '' Document
        '' =====================================================
        'ExtractContent.Run()
        'PageNumbersOfNodes.Run()
        'RemoveBreaks.Run()
        'CloningDocument.Run()
        'ProtectDocument.Run()
        'AccessStyles.Run()
        'GetVariables.Run()
        'SetViewOption.Run()
        'CreateHeaderFooterUsingDocBuilder.Run()
        'ExtractContentUsingDocumentVisitor.Run()
        'RemoveFooters.Run()
        'AddGroupShapeToDocument.Run()
        'CompareDocument.Run()
        'DocProperties.Run()
        'AcceptAllRevisions.Run()
        'WriteAndFont.Run()
        'DocumentBuilderInsertParagraph.Run()
        'DocumentBuilderBuildTable.Run()
        'DocumentBuilderInsertBreak.Run()
        'DocumentBuilderInsertImage.Run()
        'DocumentBuilderInsertBookmark.Run()
        'DocumentBuilderInsertElements.Run()
        'DocumentBuilderSetFormatting.Run()
        'DocumentBuilderMovingCursor.Run()
        'ExtractTextOnly.Run()
        'InsertDoc.Run()
        'DocumentBuilderInsertTOC.Run()
        'DocumentBuilderInsertTCField.Run()
        'DocumentBuilderInsertTCFieldsAtText.Run()
        'RemoveTOCFromDocument.Run()
        'CheckBoxTypeContentControl.Run()
        'RichTextBoxContentControl.Run()
        'ComboBoxContentControl.Run()
        'UpdateContentControls.Run()

        '' Fields
        '' =====================================================
        'InsertNestedFields.Run()
        'RemoveField.Run()
        'ConvertFieldsInDocument.Run()
        'ConvertFieldsInBody.Run()
        'ConvertFieldsInParagraph.Run()
        'ChangeLocale.Run()
        'UpdateDocFields.Run()
        'InsertField.Run()
        'InsertMergeFieldUsingDOM.Run()
        'InsertMailMergeAddressBlockFieldUsingDOM.Run()
        'InsertAdvanceFieldWithOutDocumentBuilder.Run()
        'InsertASKFieldWithOutDocumentBuilder.Run()
        'InsertAuthorField.Run()
        'InsertFormFields.Run()
        'FormFieldsGetFormFieldsCollection.Run()
        'FormFieldsGetByName.Run()
        'FormFieldsWorkWithProperties.Run()
        'RenameMergeFields.Run()
        'ChangeFieldUpdateCultureSource.Run()
        'GetFieldNames.Run()

        '' Images
        '' =====================================================
        'AddImageToEachPage.Run()
        'AddWatermark.Run()
        'CompressImages.Run()
        'ExtractImagesToFiles.Run()
        'InsertBarcodeImage.Run()

        '' Styles
        '' =====================================================
        'ExtractContentBasedOnStyles.Run()
        'ChangeStyleOfTOCLevel.Run()
        'ChangeTOCTabStops.Run()

        '' Hyperlink
        '' =====================================================
        'ReplaceHyperlinks.Run()

        '' Ranges
        '' =====================================================
        'RangesGetText.Run()
        'RangesDeleteText.Run()

        '' Tables
        '' =====================================================
        'AutoFitTableToWindow.Run()
        'AutoFitTableToContents.Run()
        'AutoFitTableToFixedColumnWidths.Run()
        'InsertTableUsingDocumentBuilder.Run()
        'InsertTableFromHtml.Run()
        'InsertTableDirectly.Run()
        'CloneTable.Run()
        'ApplyFormatting.Run()
        'SpecifyHeightAndWidth.Run()
        'ApplyStyle.Run()
        'ExtractText.Run()
        'FindingIndex.Run()
        'AddRemoveColumn.Run()
        'RepeatRowsOnSubsequentPages.Run()
        'JoiningAndSplittingTable.Run() 
        'MergedCells.Run()
        'KeepTablesAndRowsBreaking.Run()

        '' Sections
        '' =====================================================
        'SectionsAccessByIndex.Run()
        'AddDeleteSection.Run()
        'AppendSectionContent.Run()
        'DeleteSectionContent.Run()
        'DeleteHeaderFooterContent.Run()
        'CloneSection.Run()
        'CopySection.Run()

        '' =====================================================
        '' =====================================================
        '' MailMerge and Reporting
        '' =====================================================
        '' =====================================================

        'ApplyCustomLogicToEmptyRegions.Run()
        'LINQtoXMLMailMerge.Run()
        'SimpleMailMerge.Run()
        'MailMergeFormFields.Run()
        'MultipleDocsInMailMerge.Run()
        'NestedMailMerge.Run()
        'RemoveEmptyRegions.Run()
        'XMLMailMerge.Run()
        'ExecuteArray.Run()
        'MailMergeAlternatingRows.Run()
        'MailMergeImageFromBlob.Run()
        'ProduceMultipleDocuments.Run()
        'MailMergeUsingMustacheSyntax.Run()
        'ExecuteWithRegionsDataTable.Run()
        'NestedMailMergeCustom.Run()

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
        'ReadActiveXControlProperties.Run()
        'SetTrueTypeFontsFolder.Run()
        'SetFontsFoldersMultipleFolders.Run()
        'SetFontsFoldersSystemAndCustomFolder.Run()
        'SpecifyDefaultFontWhenRendering.Run()
        'ReceiveNotificationsOfFont.Run()
        'EmbeddedFontsInPDF.Run()
        'EmbeddingWindowsStandardFonts.Run()
        'HyphenateWordsOfLanguages.Run()
        'LoadHyphenationDictionaryForLanguage.Run()
        'PrintProgressDialog.Run()

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
        Return Path.GetFullPath(GetDataDir_Data() + "LINQ/")
    End Function
    Public Function GetDataDir_Database() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Database/")
    End Function
    Public Function GetDataDir_WorkingWithNode() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Node/")
    End Function
    Public Function GetDataDir_LoadingAndSaving() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Loading-and-Saving/")
    End Function
    Public Function GetDataDir_ConvertUtil() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/ConvertUtil/")
    End Function
    Public Function GetDataDir_JoiningAndAppending() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Joining-Appending/")
    End Function
    Public Function GetDataDir_WorkingWithHyperlink() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Hyperlink/")
    End Function
    Public Function GetDataDir_WorkingWithCharts() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Charts/")
    End Function
    Public Function GetDataDir_FindAndReplace() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Find-Replace/")
    End Function

    Public Function GetDataDir_WorkingWithBookmarks() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Bookmarks/")
    End Function

    Public Function GetDataDir_WorkingWithComments() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Comments/")
    End Function

    Public Function GetDataDir_WorkingWithDocument() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Document/")
    End Function

    Public Function GetDataDir_WorkingWithFields() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Fields/")
    End Function
    Public Function GetDataDir_WorkingWithRanges() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Ranges/")
    End Function

    Public Function GetDataDir_WorkingWithImages() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Images/")
    End Function

    Public Function GetDataDir_WorkingWithStyles() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Styles/")
    End Function

    Public Function GetDataDir_WorkingWithTables() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Tables/")
    End Function
    Public Function GetDataDir_WorkingWithSections() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Sections/")
    End Function

    Public Function GetDataDir_MailMergeAndReporting() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Mail-Merge/")
    End Function

    Public Function GetDataDir_QuickStart() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Quick-Start/")
    End Function
    Public Function GetDataDir_WorkingWithTheme() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Theme/")
    End Function

    Public Function GetDataDir_RenderingAndPrinting() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Rendering-Printing/")
    End Function

    Public Function GetDataDir_ViewersAndVisualizers() As [String]
        Return Path.GetFullPath(GetDataDir_Data() + "Viewers-Visualizers/")
    End Function
    Private Function GetDataDir_Data() As String
        Dim parent = Directory.GetParent(Directory.GetCurrentDirectory()).Parent
        Dim startDirectory As String = Nothing
        If parent IsNot Nothing Then
            Dim directoryInfo = parent.Parent
            If directoryInfo IsNot Nothing Then
                startDirectory = directoryInfo.FullName
            End If
        Else
            startDirectory = parent.FullName
        End If
        Return Path.Combine(startDirectory, "Data\")
    End Function
    Public Function GetOutputFilePath(inputFilePath As [String]) As String
        Dim extension As String = Path.GetExtension(inputFilePath)
        Dim filename As String = Path.GetFileNameWithoutExtension(inputFilePath)
        Return Convert.ToString(filename & Convert.ToString("_out_")) & extension
    End Function

End Module
