Imports System.Collections
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Tables
Imports Aspose.Words.Fields

Class ExtractTextOnly
    Public Shared Sub Run()
        ' ExStart:ExtractTextOnly
        Dim doc As New Document()

        ' Enter a dummy field into the document.
        Dim builder As New DocumentBuilder(doc)
        builder.InsertField("MERGEFIELD Field")

        ' GetText will retrieve all field codes and special characters
        Console.WriteLine("GetText() Result: " + doc.GetText())

        ' ToString will export the node to the specified format. When converted to text it will not retrieve fields code 
        ' or special characters, but will still contain some natural formatting characters such as paragraph markers etc. 
        ' This is the same as "viewing" the document as if it was opened in a text editor.
        Console.WriteLine("ToString() Result: " + doc.ToString(SaveFormat.Text))
        ' ExEnd:ExtractTextOnly            
    End Sub
End Class
