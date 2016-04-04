' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()

' Open the document we want to convert.
Dim doc As New Document(dataDir & Convert.ToString("Visitor.ToText.doc"))

' Create an object that inherits from the DocumentVisitor class.
Dim myConverter As New MyDocToTxtWriter()

' This is the well known Visitor pattern. Get the model to accept a visitor.
' The model will iterate through itself by calling the corresponding methods
' on the visitor object (this is called visiting).
' 
' Note that every node in the object model has the Accept method so the visiting
' can be executed not only for the whole document, but for any node in the document.
doc.Accept(myConverter)

' Once the visiting is complete, we can retrieve the result of the operation,
' that in this example, has accumulated in the visitor.
Console.WriteLine(myConverter.GetText())
