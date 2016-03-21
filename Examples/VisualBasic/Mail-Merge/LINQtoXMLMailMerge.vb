Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.IO
Imports System.Reflection
Imports Aspose.Words.MailMerging

#If (Not NET20) Then
Imports System.Linq
Imports System.Xml.Linq
#End If

Imports Aspose.Words
Imports Aspose.Words.Reporting

Public Class LINQtoXMLMailMerge
    Public Shared Sub Run()
#If (Not NET20) Then
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()

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
        ' ExStart:LINQtoXMLMailMergeorderItems
        Dim orderItems = From order In orderXml.Descendants("Item") _
        Select New With {Key .PartNumber = CStr(order.Attribute("PartNumber")), Key .ProductName = CStr(order.Element("ProductName")), Key .Quantity = CStr(order.Element("Quantity")), Key .USPrice = CStr(order.Element("USPrice")), Key .Comment = CStr(order.Element("Comment")), Key .ShipDate = CStr(order.Element("ShipDate"))}
        ' ExEnd:LINQtoXMLMailMergeorderItems
        ' ExStart:LINQToXMLQueryForDeliveryAddress
        ' Query the delivery (shipping) address using LINQ.
        Dim deliveryAddress = From delivery In orderXml.Elements("Address") _
        Where (CStr(delivery.Attribute("Type")) = "Shipping") _
        '                        Select New With {Key .Name = CStr(delivery.Element("Name")), Key .Country = CStr(delivery.Element("Country")), Key .Zip = CStr(delivery.Element("Zip")), Key .State = CStr(delivery.Element("State")), Key .City = CStr(delivery.Element("City")), Key .Street = CStr(delivery.Element("Street"))}
        ' ExEnd:LINQToXMLQueryForDeliveryAddress
        
        ' Create custom Aspose.Words mail merge data sources based on the LINQ queries.
        Dim orderItemsDataSource As New MyMailMergeDataSource(orderItems, "Items")
        Dim deliveryDataSource As New MyMailMergeDataSource(deliveryAddress)
        ' ExStart:LINQToXMLMailMerge
        Dim fileName As String = "TestFile.doc"
        ' Open the template document.
        Dim doc As New Document(dataDir & fileName)

        ' Fill the document with data from our data sources.
        ' Using mail merge regions for populating the order items table is required
        ' because it allows the region to be repeated in the document for each order item.
        doc.MailMerge.ExecuteWithRegions(orderItemsDataSource)

        ' The standard mail merge without regions is used for the delivery address.
        doc.MailMerge.Execute(deliveryDataSource)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Save the output document.
        doc.Save(dataDir)
        ' ExEnd:LINQToXMLMailMerge
        Console.WriteLine(vbNewLine & "Mail merge performed successfully." & vbNewLine & "File saved at " + dataDir)
#Else
            Throw New InvalidOperationException("This example requires the .NET Framework v3.5 or above to run." & " Make sure that the target framework of this project is set to 3.5 or above.")
#End If
    End Sub
    ' ExStart:MyMailMergeDataSource
    Public Class MyMailMergeDataSource
        Implements IMailMergeDataSource
        ' ExEnd:MyMailMergeDataSource
        ' ExStart:MyMailMergeDataSourceConstructor 
        Public Sub New(ByVal data As IEnumerable)
            mEnumerator = data.GetEnumerator()
        End Sub
        ' ExEnd:MyMailMergeDataSourceConstructor 
        ' ExStart:MyMailMergeDataSourceConstructorWithDataTable
        Public Sub New(ByVal data As IEnumerable, ByVal tableName As String)
            mEnumerator = data.GetEnumerator()
            mTableName = tableName
        End Sub
        ' ExEnd:MyMailMergeDataSourceConstructorWithDataTable
        ' ExStart:MyMailMergeDataSourceGetValue
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
        ' ExStart:MyMailMergeDataSourceGetValue
        ' ExStart:MyMailMergeDataSourceMoveNext
        Public Function MoveNext() As Boolean Implements IMailMergeDataSource.MoveNext
            Return mEnumerator.MoveNext()
        End Function
        ' ExEnd:MyMailMergeDataSourceMoveNext
        ' ExStart:MyMailMergeDataSourceTableName
        Public ReadOnly Property TableName() As String Implements IMailMergeDataSource.TableName
            Get
                Return mTableName
            End Get
        End Property
        ' ExEnd:MyMailMergeDataSourceTableName
        Public Function GetChildDataSource(ByVal tableName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
            Return Nothing
        End Function

        Private ReadOnly mEnumerator As IEnumerator
        Private ReadOnly mTableName As String
    End Class

End Class
