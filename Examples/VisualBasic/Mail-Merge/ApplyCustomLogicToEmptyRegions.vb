
Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Data
Imports System.Reflection
Imports System.Collections

Imports Aspose.Words
Imports Aspose.Words.Reporting
Imports Aspose.Words.Tables
Imports Aspose.Words.Replacing
Imports Aspose.Words.MailMerging

Public Class ApplyCustomLogicToEmptyRegions
    Public Shared Sub Run()
        ' ExStart:ApplyCustomLogicToEmptyRegions
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()
        Dim fileName As String = "TestFile.doc"
        ' Open the document.
        Dim doc As New Document(dataDir & fileName)

        ' Create a data source which has some data missing.
        ' This will result in some regions that are merged and some that remain after executing mail merge.
        Dim data As DataSet = GetDataSource()

        ' Make sure that we have not set the removal of any unused regions as we will handle them manually.
        ' We achieve this by removing the RemoveUnusedRegions flag from the cleanup options by using the AND and NOT bitwise operators.
        doc.MailMerge.CleanupOptions = doc.MailMerge.CleanupOptions And Not MailMergeCleanupOptions.RemoveUnusedRegions

        ' Execute mail merge. Some regions will be merged with data, others left unmerged.
        doc.MailMerge.ExecuteWithRegions(data)

        ' The regions which contained data now would of been merged. Any regions which had no data and were
        ' not merged will still remain in the document.
        Dim mergedDoc As Document = doc.Clone() 'ExSkip
        ' Apply logic to each unused region left in the document using the logic set out in the handler.
        ' The handler class must implement the IFieldMergingCallback interface.
        ExecuteCustomLogicOnEmptyRegions(doc, New EmptyRegionsHandler())

        ' Save the output document to disk.
        doc.Save(dataDir & "TestFile.CustomLogicEmptyRegions1_out_.doc")
        ' ExEnd:ApplyCustomLogicToEmptyRegions
        ' Reload the original merged document.
        doc = mergedDoc.Clone()

        ' Apply different logic to unused regions this time.
        ExecuteCustomLogicOnEmptyRegions(doc, New EmptyRegionsHandler_MergeTable())

        doc.Save(dataDir & "TestFile.CustomLogicEmptyRegions2_out_.doc")

        ' Reload the original merged document.
        doc = mergedDoc.Clone()
        ' ExStart:ContactDetails
        Dim regions As New ArrayList()
        regions.Add("ContactDetails")
        ExecuteCustomLogicOnEmptyRegions(doc, New EmptyRegionsHandler(), regions)
        ' ExEnd:ContactDetails
        doc.Save(dataDir & "TestFile.CustomLogicEmptyRegions3_out_.doc")

        Console.WriteLine(vbNewLine + "Mail merge performed successfully." + vbNewLine + "File saved at " + dataDir + "TestFile.CustomLogicEmptyRegions3_out_.doc")
    End Sub
    ' ExStart:CreateDataSourceFromDocumentRegions
    Private Shared Function CreateDataSourceFromDocumentRegions(ByVal doc As Document, ByVal regionsList As ArrayList) As DataSet
        Const tableStartMarker As String = "TableStart:"
        Dim dataSet As New DataSet()
        Dim tableName As String = Nothing

        For Each fieldName As String In doc.MailMerge.GetFieldNames()
            If fieldName.Contains(tableStartMarker) Then
                tableName = fieldName.Substring(tableStartMarker.Length)
            ElseIf tableName IsNot Nothing Then
                ' Only add the table name as a new DataTable if it doesn't already exists in the DataSet.
                If dataSet.Tables(tableName) Is Nothing Then
                    Dim table As New DataTable(tableName)
                    table.Columns.Add(fieldName)

                    ' We only need to add the first field for the handler to be called for the fields in the region.
                    If regionsList Is Nothing OrElse regionsList.Contains(tableName) Then
                        table.Rows.Add("FirstField")
                    End If

                    dataSet.Tables.Add(table)
                End If
                tableName = Nothing
            End If
        Next fieldName

        Return dataSet
    End Function
    ' ExEnd:CreateDataSourceFromDocumentRegions
    ' ExStart:ExecuteCustomLogicOnEmptyRegions
    Public Shared Sub ExecuteCustomLogicOnEmptyRegions(ByVal doc As Document, ByVal handler As IFieldMergingCallback)
        ExecuteCustomLogicOnEmptyRegions(doc, handler, Nothing) ' Pass null to handle all regions found in the document.
    End Sub

    Public Shared Sub ExecuteCustomLogicOnEmptyRegions(ByVal doc As Document, ByVal handler As IFieldMergingCallback, ByVal regionsList As ArrayList)
        ' Certain regions can be skipped from applying logic to by not adding the table name inside the CreateEmptyDataSource method.
        ' Enable this cleanup option so any regions which are not handled by the user's logic are removed automatically.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions

        ' Set the user's handler which is called for each unmerged region.
        doc.MailMerge.FieldMergingCallback = handler

        ' Execute mail merge using the dummy dataset. The dummy data source contains the table names of 
        ' each unmerged region in the document (excluding ones that the user may have specified to be skipped). This will allow the handler 
        ' to be called for each field in the unmerged regions.
        doc.MailMerge.ExecuteWithRegions(CreateDataSourceFromDocumentRegions(doc, regionsList))
    End Sub
    ' ExEnd:ExecuteCustomLogicOnEmptyRegions
    ' ExStart:EmptyRegionsHandler
    Public Class EmptyRegionsHandler
        Implements IFieldMergingCallback
        ''' <summary>
        ''' Called for each field belonging to an unmerged region in the document.
        ''' </summary>
        Public Sub FieldMerging(ByVal args As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
            ' Change the text of each field of the ContactDetails region individually.
            If args.TableName = "ContactDetails" Then
                ' Set the text of the field based off the field name.
                If args.FieldName = "Name" Then
                    args.Text = "(No details found)"
                ElseIf args.FieldName = "Number" Then
                    args.Text = "(N/A)"
                End If
            End If

            ' Remove the entire table of the Suppliers region. Also check if the previous paragraph
            ' before the table is a heading paragraph and if so remove that too.
            If args.TableName = "Suppliers" Then
                Dim table As Table = CType(args.Field.Start.GetAncestor(NodeType.Table), Table)

                ' Check if the table has been removed from the document already.
                If table.ParentNode IsNot Nothing Then
                    ' Try to find the paragraph which precedes the table before the table is removed from the document.
                    If table.PreviousSibling IsNot Nothing AndAlso table.PreviousSibling.NodeType = NodeType.Paragraph Then
                        Dim previousPara As Paragraph = CType(table.PreviousSibling, Paragraph)
                        If IsHeadingParagraph(previousPara) Then
                            previousPara.Remove()
                        End If
                    End If

                    table.Remove()
                End If
            End If
        End Sub

        ''' <summary>
        ''' Returns true if the paragraph uses any Heading style e.g Heading 1 to Heading 9
        ''' </summary>
        Private Function IsHeadingParagraph(ByVal para As Paragraph) As Boolean
            Return (para.ParagraphFormat.StyleIdentifier >= StyleIdentifier.Heading1 AndAlso para.ParagraphFormat.StyleIdentifier <= StyleIdentifier.Heading9)
        End Function

        Public Sub ImageFieldMerging(ByVal args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
            ' Do Nothing
        End Sub
    End Class
    ' ExEnd:EmptyRegionsHandler
    Public Class EmptyRegionsHandler_MergeTable
        Implements IFieldMergingCallback
        ''' <summary>
        ''' Called for each field belonging to an unmerged region in the document.
        ''' </summary>
        Public Sub FieldMerging(ByVal args As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
            ' ExStart:RemoveExtraParagraphs
            ' Store the parent paragraph of the current field for easy access.
            Dim parentParagraph As Paragraph = args.Field.Start.ParentParagraph

            ' Define the logic to be used when the ContactDetails region is encountered.
            ' The region is removed and replaced with a single line of text stating that there are no records.
            If args.TableName = "ContactDetails" Then
                ' Called for the first field encountered in a region. This can be used to execute logic on the first field
                ' in the region without needing to hard code the field name. Often the base logic is applied to the first field and 
                ' different logic for other fields. The rest of the fields in the region will have a null FieldValue.
                If CStr(args.FieldValue) Is "FirstField" Then
                    ' Remove the "Name:" tag from the start of the paragraph
                    parentParagraph.Range.Replace("Name:", String.Empty, New FindReplaceOptions(FindReplaceDirection.Forward))
                    ' Set the text of the first field to display a message stating that there are no records.
                    args.Text = "No records to display"
                Else
                    ' We have already inserted our message in the paragraph belonging to the first field. The other paragraphs in the region
                    ' will still remain so we want to remove these. A check is added to ensure that the paragraph has not already been removed.
                    ' which may happen if more than one field is included in a paragraph.
                    If parentParagraph.ParentNode IsNot Nothing Then
                        parentParagraph.Remove()
                    End If
                End If
            End If
            ' ExEnd:RemoveExtraParagraphs
            ' ExStart:MergeAllCells
            ' Replace the unused region in the table with a "no records" message and merge all cells into one.
            If args.TableName = "Suppliers" Then
                If CStr(args.FieldValue) Is "FirstField" Then
                    ' We will use the first paragraph to display our message. Make it centered within the table. The other fields in other cells 
                    ' within the table will be merged and won't be displayed so we don't need to do anything else with them.
                    parentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center
                    args.Text = "No records to display"
                End If

                ' Merge the cells of the table together. 
                Dim cell As Cell = CType(parentParagraph.GetAncestor(NodeType.Cell), Cell)
                If cell IsNot Nothing Then
                    If cell.IsFirstCell Then
                        cell.CellFormat.HorizontalMerge = CellMerge.First ' If this cell is the first cell in the table then the merge is started using "CellMerge.First".
                    Else
                        cell.CellFormat.HorizontalMerge = CellMerge.Previous ' Otherwise the merge is continued using "CellMerge.Previous".
                    End If
                End If
            End If
            ' ExEnd:MergeAllCells
        End Sub

        Public Sub ImageFieldMerging(ByVal args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
            ' Do Nothing
        End Sub

        
    End Class

    Private Shared Function GetDataSource() As DataSet
        ' Create a new DataSet and DataTable objects to be used for mail merge.
        Dim data As New DataSet()
        Dim storeDetails As New DataTable("StoreDetails")
        Dim contactDetails As New DataTable("ContactDetails")

        ' Add columns for the ContactDetails table.
        contactDetails.Columns.Add("ID")
        contactDetails.Columns.Add("Name")
        contactDetails.Columns.Add("Number")

        ' Add columns for the StoreDetails table.
        storeDetails.Columns.Add("ID")
        storeDetails.Columns.Add("Name")
        storeDetails.Columns.Add("Address")
        storeDetails.Columns.Add("City")
        storeDetails.Columns.Add("Country")

        ' Add the data to the tables.
        storeDetails.Rows.Add("0", "Hungry Coyote Import Store", "2732 Baker Blvd", "Eugene", "USA")
        storeDetails.Rows.Add("1", "Great Lakes Food Market", "City Center Plaza, 516 Main St.", "San Francisco", "USA")

        ' Add data to the child table only for the first record.
        contactDetails.Rows.Add("0", "Thomas Hardy", "(206) 555-9857 ext 237")
        contactDetails.Rows.Add("0", "Elizabeth Brown", "(206) 555-9857 ext 764")

        ' Include the tables in the DataSet.
        data.Tables.Add(storeDetails)
        data.Tables.Add(contactDetails)

        ' Setup the relation between the parent table (StoreDetails) and the child table (ContactDetails).
        data.Relations.Add(storeDetails.Columns("ID"), contactDetails.Columns("ID"))

        Return data
    End Function
    Private Shared orderTable As DataTable = Nothing
    Private Shared itemTable As DataTable = Nothing
    Private Shared Sub DisableForeignKeyConstraints(dataSet As DataSet)
        ' ExStart:DisableForeignKeyConstraints
        dataSet.Relations.Add(New DataRelation("OrderToItem", orderTable.Columns("Order_Id"), itemTable.Columns("Order_Id"), False))
        ' ExEnd:DisableForeignKeyConstraints
    End Sub
    Private Shared Sub CreateDataRelation(dataSet As DataSet)
        ' ExStart:CreateDataRelation
        dataSet.Relations.Add(New DataRelation("OrderToItem", orderTable.Columns("Order_Id"), itemTable.Columns("Order_Id")))
        ' ExEnd:CreateDataRelation
    End Sub
End Class
