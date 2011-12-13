'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Linq
Imports System.Xml.Linq
Imports System.IO
Imports System.Reflection

Imports Aspose.Words
Imports Aspose.Words.Reporting


Namespace LINQtoXMLMailMerge
	Friend Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

			' Load the XML document.
			Dim orderXml As XElement = XElement.Load(dataDir & "PurchaseOrder.xml")

			' Query the purchase order xml file using LINQ to extract the order items 
			' into an object of an anonymous type. 
			'
			' Make sure you give the properties of the anonymous type the same names as 
			' the MERGEFIELD fields in the document.
			'
			' To pass the actual values stored in the XML element or attribute to Aspose.Words, 
			' we need to cast them to string. This is to prevent the XML tags being inserted into the final document when
			' the XElement or XAttribute objects are passed to Aspose.Words.

			'ExStart
			'ExId:LINQtoXMLMailMerge_query_items
			'ExSummary:LINQ to XML query for ordered items.
			Dim orderItems = From order In orderXml.Descendants("Item") _
			                 Select New With {Key .PartNumber = CStr(order.Attribute("PartNumber")), Key .ProductName = CStr(order.Element("ProductName")), Key .Quantity = CStr(order.Element("Quantity")), Key .USPrice = CStr(order.Element("USPrice")), Key .Comment = CStr(order.Element("Comment")), Key .ShipDate = CStr(order.Element("ShipDate"))}
			'ExEnd

			' Query the delivery (shipping) address using LINQ.
			'ExStart
			'ExId:LINQtoXMLMailMerge_query_delivery
			'ExSummary:LINQ to XML query for delivery address.
			Dim deliveryAddress = From delivery In orderXml.Elements("Address") _
			                      Where (CStr(delivery.Attribute("Type")) = "Shipping") _
			                      Select New With {Key .Name = CStr(delivery.Element("Name")), Key .Country = CStr(delivery.Element("Country")), Key .Zip = CStr(delivery.Element("Zip")), Key .State = CStr(delivery.Element("State")), Key .City = CStr(delivery.Element("City")), Key .Street = CStr(delivery.Element("Street"))}
			'ExEnd

			' Create custom Aspose.Words mail merge data sources based on the LINQ queries.
			Dim orderItemsDataSource As New MyMailMergeDataSource(orderItems, "Items")
			Dim deliveryDataSource As New MyMailMergeDataSource(deliveryAddress)

			'ExStart
			'ExFor:MailMerge.ExecuteWithRegions(Aspose.Words.Reporting.IMailMergeDataSource)
			'ExId:LINQtoXMLMailMerge_call
			'ExSummary:Perform the mail merge and save the result.
			' Open the template document.
			Dim doc As New Document(dataDir & "TestFile.doc")

			' Fill the document with data from our data sources.
			' Using mail merge regions for populating the order items table is required
			' because it allows the region to be repeated in the document for each order item.
			doc.MailMerge.ExecuteWithRegions(orderItemsDataSource)

			' The standard mail merge without regions is used for the delivery address.
			doc.MailMerge.Execute(deliveryDataSource)

			' Save the output document.
			doc.Save(dataDir & "TestFile Out.doc")
			'ExEnd
		End Sub

		''' <summary>
		''' Aspose.Words does not accept LINQ queries as an input for mail merge directly, 
		''' but provides a generic mechanism which allows mail merges from any data source.
		''' 
		''' This class is a simple implementation of the Aspose.Words custom mail merge data source 
		''' interface that accepts a LINQ query (in fact any IEnumerable object).
		''' Aspose.Words calls this class during the mail merge to retrieve the data.
		''' </summary>
		'ExStart
		'ExId:LINQtoXMLMailMerge_class
		'ExSummary:The implementation of the IMailMergeDataSource interface.
		Public Class MyMailMergeDataSource
			Implements IMailMergeDataSource
		'ExEnd
			''' <summary>
			''' Creates a new instance of a custom mail merge data source.
			''' </summary>
			''' <param name="data">Data returned from a LINQ query.</param>
			'ExStart
			'ExId:LINQtoXMLMailMerge_constructor_simple
			'ExSummary:Constructor for the simple mail merge.
			Public Sub New(ByVal data As IEnumerable)
				mEnumerator = data.GetEnumerator()
			End Sub
			'ExEnd

			''' <summary>
			''' Creates a new instance of a custom mail merge data source, for mail merge with regions.
			''' </summary>
			''' <param name="data">Data returned from a LINQ query.</param>
			''' <param name="tableName">Name of the data source is only used when you perform mail merge with regions. 
			''' If you prefer to use the simple mail merge then use constructor with one parameter.</param>
			'ExStart
			'ExId:LINQtoXMLMailMerge_constructor_with_regions
			'ExSummary:Constructor for the mail merge with regions.
			Public Sub New(ByVal data As IEnumerable, ByVal tableName As String)
				mEnumerator = data.GetEnumerator()
				mTableName = tableName
			End Sub
			'ExEnd

			''' <summary>
			''' Aspose.Words calls this method to get a value for every data field.
			''' 
			''' This is a simple "generic" implementation of a data source that can work over 
			''' any IEnumerable collection. This implementation assumes that the merge field
			''' name in the document matches the name of a public property on the object
			''' in the collection and uses reflection to get the value of the property.
			''' </summary>
			'ExStart
			'ExId:LINQtoXMLMailMerge_get_value
			'ExSummary:Getting the field value in the custom data source.
			Public Function GetValue(ByVal fieldName As String, <System.Runtime.InteropServices.Out()> ByRef fieldValue As Object) As Boolean Implements IMailMergeDataSource.GetValue
				' Use reflection to get the property by name from the current object.
				Dim obj As Object = mEnumerator.Current

				Dim curentRecordType As Type = obj.GetType()
				Dim [property] As PropertyInfo = curentRecordType.GetProperty(fieldName)
				If [property] IsNot Nothing Then
					fieldValue = [property].GetValue(obj, Nothing)
					Return True
				End If

				' Return False to the Aspose.Words mail merge engine to indicate the field was not found.
				fieldValue = Nothing
				Return False
			End Function
			'ExEnd

			''' <summary>
			''' Moves to the next record in the collection.
			''' </summary>
			'ExStart
			'ExId:LINQtoXMLMailMerge_move_next
			'ExSummary:Moving through the data records.
			Public Function MoveNext() As Boolean Implements IMailMergeDataSource.MoveNext
				Return mEnumerator.MoveNext()
			End Function
			'ExEnd

			''' <summary>
			''' The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
			''' </summary>
			'ExStart
			'ExId:LINQtoXMLMailMerge_table_name
			'ExSummary:The table name property.
			Public ReadOnly Property TableName() As String Implements IMailMergeDataSource.TableName
				Get
					Return mTableName
				End Get
			End Property
			'ExEnd

			Public Function GetChildDataSource(ByVal tableName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
				Return Nothing
			End Function

			Private ReadOnly mEnumerator As IEnumerator
			Private ReadOnly mTableName As String
		End Class
	End Class
End Namespace
