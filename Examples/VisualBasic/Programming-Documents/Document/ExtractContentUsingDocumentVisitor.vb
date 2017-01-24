Imports System.IO
Imports Aspose.Words
Imports System
Imports System.Text
Imports Aspose.Words.Fields
Public Class ExtractContentUsingDocumentVisitor
    Public Shared Sub Run()
        ' ExStart:ExtractContentUsingDocumentVisitor
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()

        ' Open the document we want to convert.
        Dim doc As New Document(dataDir & Convert.ToString("Visitor.ToText.doc"))

        ' Create an object that inherits from the DocumentVisitor class.
        Dim myConverter As New MyDocToTxtWriter()

        ' This is the well known Visitor pattern. Get the model to accept a visitor.
        ' The model will iterate through itself by calling the corresponding methods
        ' On the visitor object (this is called visiting).
        ' 
        ' Note that every node in the object model has the Accept method so the visiting
        ' Can be executed not only for the whole document, but for any node in the document.
        doc.Accept(myConverter)

        ' Once the visiting is complete, we can retrieve the result of the operation,
        ' That in this example, has accumulated in the visitor.
        Console.WriteLine(myConverter.GetText())
        ' ExEnd:ExtractContentUsingDocumentVisitor
    End Sub
    ' ExStart:MyDocToTxtWriter
    ''' <summary>
    ''' Simple implementation of saving a document in the plain text format. Implemented as a Visitor.
    ''' </summary>
    Friend Class MyDocToTxtWriter
        Inherits DocumentVisitor
        Public Sub New()
            mIsSkipText = False
            mBuilder = New StringBuilder()
        End Sub

        ''' <summary>
        ''' Gets the plain text of the document that was accumulated by the visitor.
        ''' </summary>
        Public Function GetText() As String
            Return mBuilder.ToString()
        End Function

        ''' <summary>
        ''' Called when a Run node is encountered in the document.
        ''' </summary>
        Public Overrides Function VisitRun(run As Run) As VisitorAction
            AppendText(run.Text)

            ' Let the visitor continue visiting other nodes.
            Return VisitorAction.[Continue]
        End Function

        ''' <summary>
        ''' Called when a FieldStart node is encountered in the document.
        ''' </summary>
        Public Overrides Function VisitFieldStart(fieldStart As FieldStart) As VisitorAction
            ' In Microsoft Word, a field code (such as "MERGEFIELD FieldName") follows
            ' After a field start character. We want to skip field codes and output field 
            ' Result only, therefore we use a flag to suspend the output while inside a field code.
            '
            ' Note this is a very simplistic implementation and will not work very well
            ' If you have nested fields in a document. 
            mIsSkipText = True

            Return VisitorAction.[Continue]
        End Function

        ''' <summary>
        ''' Called when a FieldSeparator node is encountered in the document.
        ''' </summary>
        Public Overrides Function VisitFieldSeparator(fieldSeparator As FieldSeparator) As VisitorAction
            ' Once reached a field separator node, we enable the output because we are
            ' Now entering the field result nodes.
            mIsSkipText = False

            Return VisitorAction.[Continue]
        End Function

        ''' <summary>
        ''' Called when a FieldEnd node is encountered in the document.
        ''' </summary>
        Public Overrides Function VisitFieldEnd(fieldEnd As FieldEnd) As VisitorAction
            ' Make sure we enable the output when reached a field end because some fields
            ' Do not have field separator and do not have field result.
            mIsSkipText = False

            Return VisitorAction.[Continue]
        End Function

        ''' <summary>
        ''' Called when visiting of a Paragraph node is ended in the document.
        ''' </summary>
        Public Overrides Function VisitParagraphEnd(paragraph As Paragraph) As VisitorAction
            ' When outputting to plain text we output Cr+Lf characters.
            AppendText(ControlChar.CrLf)

            Return VisitorAction.[Continue]
        End Function

        Public Overrides Function VisitBodyStart(body As Body) As VisitorAction
            ' We can detect beginning and end of all composite nodes such as Section, Body, 
            ' Table, Paragraph etc and provide custom handling for them.
            mBuilder.Append("*** Body Started ***" & vbCr & vbLf)

            Return VisitorAction.[Continue]
        End Function

        Public Overrides Function VisitBodyEnd(body As Body) As VisitorAction
            mBuilder.Append("*** Body Ended ***" & vbCr & vbLf)
            Return VisitorAction.[Continue]
        End Function

        ''' <summary>
        ''' Called when a HeaderFooter node is encountered in the document.
        ''' </summary>
        Public Overrides Function VisitHeaderFooterStart(headerFooter As HeaderFooter) As VisitorAction
            ' Returning this value from a visitor method causes visiting of this
            ' Node to stop and move on to visiting the next sibling node.
            ' The net effect in this example is that the text of headers and footers
            ' Is not included in the resulting output.
            Return VisitorAction.SkipThisNode
        End Function


        ''' <summary>
        ''' Adds text to the current output. Honours the enabled/disabled output flag.
        ''' </summary>
        Private Sub AppendText(text As String)
            If Not mIsSkipText Then
                mBuilder.Append(text)
            End If
        End Sub

        Private ReadOnly mBuilder As StringBuilder
        Private mIsSkipText As Boolean
    End Class
    ' ExEnd:MyDocToTxtWriter
End Class

