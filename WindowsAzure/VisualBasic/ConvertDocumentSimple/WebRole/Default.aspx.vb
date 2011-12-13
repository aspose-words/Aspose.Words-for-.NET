'ExStart
'ExId:Azure_ConvertDocumentSimple
'ExSummary:Shows how to convert a document in Windows Azure.

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.IO
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Aspose.Words

Namespace WebRole
	''' <summary>
	''' This demo shows how to use Aspose.Words for .NET inside a WebRole in a simple
	''' Windows Azure application. There is just one ASP.NET page that provides a user
	''' interface to convert a document from one format to another.
	''' </summary>
	Partial Public Class _Default
		Inherits System.Web.UI.Page
		Protected Sub SubmitButton_Click(ByVal sender As Object, ByVal e As EventArgs)
			Dim postedFile As HttpPostedFile = SrcFileUpload.PostedFile

			If postedFile.ContentLength = 0 Then
				Throw New Exception("There was no document uploaded.")
			End If

			If postedFile.ContentLength > 512 * 1024 Then
				Throw New Exception("The uploaded document is too big. This demo limits the file size to 512Kb.")
			End If

			' Create a suitable file name for the converted document.
			Dim dstExtension As String = DstFormatDropDownList.SelectedValue
			Dim dstFileName As String = Path.GetFileName(postedFile.FileName) & "_Converted." & dstExtension
			Dim dstFormat As SaveFormat = FileFormatUtil.ExtensionToSaveFormat(dstExtension)

			' Convert the document and send to the browser.
			Dim doc As New Document(postedFile.InputStream)
			doc.Save(dstFileName, dstFormat, SaveType.OpenInApplication, Response)

			' Required. Otherwise DOCX cannot be opened on the client (probably not all data sent
			' or some extra data sent in the response).
			Response.End()
		End Sub

		Shared Sub New()
			' Uncomment this code and embed your license file as a resource in this project and this code 
			' will find and activate the license. Aspose.Wods licensing needs to execute only once
			' before any Document instance is created and a static ctor is a good place.
			'
			' Aspose.Words.License l = new Aspose.Words.License();
			' l.SetLicense("Aspose.Words.lic");
		End Sub
	End Class
End Namespace
'ExEnd