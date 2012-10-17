'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports Aspose.Words
Imports Aspose.Words.Tables
Imports Aspose.Words.Drawing

Imports System
Imports System.Text
Imports System.Diagnostics
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Data.OleDb
Imports System.Reflection


Namespace ImportTableFromDataTable
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath
			' This is the location to our database. You must have the Examples folder extracted as well for the database to be found.
			Dim databaseDir As String = New Uri(New Uri(exeDir), "../../../../Examples/Common/Database/").LocalPath

			' Create the output directory if it doesn't exist.
			If (Not Directory.Exists(dataDir)) Then
				Directory.CreateDirectory(dataDir)
			End If

			'ExStart
			'ExFor:Table.StyleIdentifier
			'ExFor:StyleIdentifier
			'ExFor:Table.StyleOptions
			'ExFor:TableStyleOptions
			'ExId:ImportDataTableCaller
			'ExSummary:Shows how to import the data from a DataTable and insert it into a new table in the document.
			' Create a new document.
			Dim doc As New Document()

			' We can position where we want the table to be inserted and also specify any extra formatting to be
			' applied onto the table as well.
			Dim builder As New DocumentBuilder(doc)

			' We want to rotate the page landscape as we expect a wide table.
			doc.FirstSection.PageSetup.Orientation = Orientation.Landscape

			' Retrieve the data from our data source which is stored as a DataTable.
			Dim dataTable As DataTable = GetEmployees(databaseDir)

			' Build a table in the document from the data contained in the DataTable.
			Dim table As Table = ImportTableFromDataTable(builder, dataTable, True)

			' We can apply a table style as a very quick way to apply formatting to the entire table.
			table.StyleIdentifier = StyleIdentifier.MediumList2Accent1
			table.StyleOptions = TableStyleOptions.FirstRow Or TableStyleOptions.RowBands Or TableStyleOptions.LastColumn

			' For our table we want to remove the heading for the image column.
			table.FirstRow.LastCell.RemoveAllChildren()

			doc.Save(dataDir & "Table.FromDataTable Out.docx")
			'ExEnd

			' Do some verification on the generated table.
			doc.ExpandTableStylesToDirectFormatting()
			Debug.Assert(table.Rows.Count = 6, "Unexpected row count")
			Debug.Assert(doc.GetChildNodes(NodeType.Table, True).Count = 1, "Unexpected table count")
			Debug.Assert(table.FirstRow.FirstCell.ToString(SaveFormat.Text).Trim() = "EmployeeID", "Unexpected header text")
			Debug.Assert(table.Rows(2).Cells(2).ToString(SaveFormat.Text).Trim() = "Andrew", "Unexpected row text")
			Debug.Assert(table.Rows(1).FirstCell.CellFormat.Shading.BackgroundPatternColor <> Color.Empty, "Unexpected cell shading")
		End Sub

		'ExStart
		'ExId:ImportTableFromDataTable
		'ExSummary:Provides a method to import data from the DataTable and insert it into a new table using the DocumentBuilder.
		''' <summary>
		''' Imports the content from the specified DataTable into a new Aspose.Words Table object. 
		''' The table is inserted at the current position of the document builder and using the current builder's formatting if any is defined.
		''' </summary>
		Public Shared Function ImportTableFromDataTable(ByVal builder As DocumentBuilder, ByVal dataTable As DataTable, ByVal importColumnHeadings As Boolean) As Table
			Dim table As Table = builder.StartTable()

			' Check if the names of the columns from the data source are to be included in a header row.
			If importColumnHeadings Then
				' Store the original values of these properties before changing them.
				Dim boldValue As Boolean = builder.Font.Bold
				Dim paragraphAlignmentValue As ParagraphAlignment = builder.ParagraphFormat.Alignment

				' Format the heading row with the appropriate properties.
				builder.Font.Bold = True
				builder.ParagraphFormat.Alignment = ParagraphAlignment.Center

				' Create a new row and insert the name of each column into the first row of the table.
				For Each column As DataColumn In dataTable.Columns
					builder.InsertCell()
					builder.Writeln(column.ColumnName)
				Next column

				builder.EndRow()

				' Restore the original formatting.
				builder.Font.Bold = boldValue
				builder.ParagraphFormat.Alignment = paragraphAlignmentValue
			End If

			For Each dataRow As DataRow In dataTable.Rows
				For Each item As Object In dataRow.ItemArray
					' Insert a new cell for each object.
					builder.InsertCell()

					Select Case item.GetType().Name
						Case "Byte[]"
							' Assume a byte array is an image. Other data types can be added here.
							builder.InsertImage(GetImageFromByteArray(CType(item, Byte())), 50, 50)
						Case "DateTime"
							' Define a custom format for dates and times.
							Dim dateTime As DateTime = CDate(item)
							builder.Write(dateTime.ToString("MMMM d, yyyy"))
						Case Else
							' By default any other item will be inserted as text.
							builder.Write(item.ToString())
					End Select

				Next item

				' After we insert all the data from the current record we can end the table row.
				builder.EndRow()
			Next dataRow

			' We have finished inserting all the data from the DataTable, we can end the table.
			builder.EndTable()

			Return table
		End Function
		'ExEnd

		''' <summary>
		''' Returns a .NET Image object from the specified byte array.
		''' </summary>
		Private Shared Function GetImageFromByteArray(ByVal imageBytes() As Byte) As Image
			' Some drivers can pick up some junk data to the start of binary storage fields.
			' This means we cannot directly read the bytes into an image, we first need
			' to skip past until we find the start of the image.
			Dim imageString As String = Encoding.ASCII.GetString(imageBytes)
			Dim index As Integer = imageString.IndexOf("BM")
			Return Image.FromStream(New MemoryStream(imageBytes, index, imageBytes.Length - index))
		End Function

		''' <summary>
		''' Retrieves employee data from an external database.
		''' </summary>
		Private Shared Function GetEmployees(ByVal databaseDir As String) As DataTable
			' Open a database connection.
			Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & databaseDir & "Northwind.mdb"
			Dim conn As New OleDbConnection(connString)
			conn.Open()

			' Create the command.
			Dim cmd As New OleDbCommand("SELECT TOP 5 EmployeeID, LastName, FirstName, Title, Birthdate, Address, City, PhotoBLOB FROM Employees", conn)

			' Fill an ADO.NET table from the command.
			Dim da As New OleDbDataAdapter(cmd)
			Dim table As New DataTable()
			da.Fill(table)

			' Close database.
			conn.Close()

			Return table
		End Function
	End Class
End Namespace