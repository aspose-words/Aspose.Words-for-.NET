Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Drawing.Charts
Imports Aspose.Words.Fields
Imports Aspose.Words.Tables


Class DocumentBuilderInsertElements
    Public Shared Sub Run()

        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        InsertTextInputFormField(dataDir)
        InsertCheckBoxFormField(dataDir)
        InsertComboBoxFormField(dataDir)
        InsertHtml(dataDir)
        InsertHyperlink(dataDir)
        InsertTableOfContents(dataDir)
        InsertOleObject(dataDir)
    End Sub
    Public Shared Sub InsertTextInputFormField(dataDir As String)
        ' ExStart:DocumentBuilderInsertTextInputFormField
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Hello", 0)
        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertTextInputFormField_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertTextInputFormField
        Console.WriteLine(Convert.ToString(vbLf & "Text input form field using DocumentBuilder inserted successfully into a document." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub InsertCheckBoxFormField(dataDir As String)
        ' ExStart:DocumentBuilderInsertCheckBoxFormField
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.InsertCheckBox("CheckBox", True, True, 0)
        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertCheckBoxFormField_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertCheckBoxFormField
        Console.WriteLine(Convert.ToString(vbLf & "Checkbox form field using DocumentBuilder inserted successfully into a document." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub InsertComboBoxFormField(dataDir As String)
        ' ExStart:DocumentBuilderInsertComboBoxFormField
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim items As String() = {"One", "Two", "Three"}
        builder.InsertComboBox("DropDown", items, 0)
        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertComboBoxFormField_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertComboBoxFormField
        Console.WriteLine(Convert.ToString(vbLf & "Combobox form field using DocumentBuilder inserted successfully into a document." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub InsertHtml(dataDir As String)
        ' ExStart:DocumentBuilderInsertHtml
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.InsertHtml("<P align='right'>Paragraph right</P>" + "<b>Implicit paragraph left</b>" + "<div align='center'>Div center</div>" + "<h1 align='left'>Heading 1 left.</h1>")
        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertHtml_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertHtml
        Console.WriteLine(Convert.ToString(vbLf & "HTML using DocumentBuilder inserted successfully into a document." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub InsertHyperlink(dataDir As String)
        ' ExStart:DocumentBuilderInsertHyperlink
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.Write("Please make sure to visit ")

        ' Specify font formatting for the hyperlink.
        builder.Font.Color = Color.Blue
        builder.Font.Underline = Underline.[Single]
        ' Insert the link.
        builder.InsertHyperlink("Aspose Website", "http://www.aspose.com", False)

        ' Revert to default formatting.
        builder.Font.ClearFormatting()

        builder.Write(" for more information.")
        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertHyperlink_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertHyperlink
        Console.WriteLine(Convert.ToString(vbLf & "Hyperlink using DocumentBuilder inserted successfully into a document." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub InsertTableOfContents(dataDir As String)
        ' ExStart:DocumentBuilderInsertTableOfContents
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Insert a table of contents at the beginning of the document.
        builder.InsertTableOfContents("\o ""1-3"" \h \z \u")

        ' Start the actual document content on the second page.
        builder.InsertBreak(BreakType.PageBreak)

        ' Build a document with complex structure by applying different heading styles thus creating TOC entries.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1

        builder.Writeln("Heading 1")

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2

        builder.Writeln("Heading 1.1")
        builder.Writeln("Heading 1.2")

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1

        builder.Writeln("Heading 2")
        builder.Writeln("Heading 3")

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2

        builder.Writeln("Heading 3.1")

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3

        builder.Writeln("Heading 3.1.1")
        builder.Writeln("Heading 3.1.2")
        builder.Writeln("Heading 3.1.3")

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2

        builder.Writeln("Heading 3.2")
        builder.Writeln("Heading 3.3")
        doc.UpdateFields()

        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertTableOfContents_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertTableOfContents
        Console.WriteLine(Convert.ToString(vbLf & "Table of contents using DocumentBuilder inserted successfully into a document." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub InsertOleObject(dataDir As String)
        ' ExStart:DocumentBuilderInsertOleObject
        Dim doc As New Aspose.Words.Document()
        Dim builder As New DocumentBuilder(doc)
        builder.InsertOleObject("http://www.aspose.com", "htmlfile", True, True, Nothing)
        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertOleObject_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertOleObject
        Console.WriteLine(Convert.ToString(vbLf & "OleObject using DocumentBuilder inserted successfully into a document." & vbLf & "File saved at ") & dataDir)
    End Sub

End Class
