Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Drawing.Charts
Imports Aspose.Words.Fields
Imports Aspose.Words.Tables
Class DocumentBuilderSetFormatting
    Public Shared Sub Run()

        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        SetFontFormatting(dataDir)
        SetParagraphFormatting(dataDir)
        SetTableCellFormatting(dataDir)
        SetMultilevelListFormatting(dataDir)
        SetPageSetupAndSectionFormatting(dataDir)
        ApplyParagraphStyle(dataDir)
        ApplyBordersAndShadingToParagraph(dataDir)
    End Sub
    Public Shared Sub SetFontFormatting(dataDir As String)
        ' ExStart:DocumentBuilderSetFontFormatting
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Set font formatting properties
        Dim font As Aspose.Words.Font = builder.Font
        font.Bold = True
        font.Color = System.Drawing.Color.DarkBlue
        font.Italic = True
        font.Name = "Arial"
        font.Size = 24
        font.Spacing = 5
        font.Underline = Underline.[Double]

        ' Output formatted text
        builder.Writeln("I'm a very nice formatted string.")
        dataDir = dataDir & Convert.ToString("DocumentBuilderSetFontFormatting_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderSetFontFormatting
        Console.WriteLine(Convert.ToString(vbLf & "Font formatting using DocumentBuilder set successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub SetParagraphFormatting(dataDir As String)
        'ExStart:DocumentBuilderSetParagraphFormatting
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Set paragraph formatting properties
        Dim paragraphFormat As ParagraphFormat = builder.ParagraphFormat
        paragraphFormat.Alignment = ParagraphAlignment.Center
        paragraphFormat.LeftIndent = 50
        paragraphFormat.RightIndent = 50
        paragraphFormat.SpaceAfter = 25

        ' Output text
        builder.Writeln("I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.")
        builder.Writeln("I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.")

        dataDir = dataDir & Convert.ToString("DocumentBuilderSetParagraphFormatting_out_.doc")
        doc.Save(dataDir)
        'ExEnd:DocumentBuilderSetParagraphFormatting
        Console.WriteLine(Convert.ToString(vbLf & "Paragraph formatting using DocumentBuilder set successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub SetTableCellFormatting(dataDir As String)
        ' ExStart:DocumentBuilderSetTableCellFormatting
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.StartTable()
        builder.InsertCell()

        ' Set the cell formatting
        Dim cellFormat As CellFormat = builder.CellFormat
        cellFormat.Width = 250
        cellFormat.LeftPadding = 30
        cellFormat.RightPadding = 30
        cellFormat.TopPadding = 30
        cellFormat.BottomPadding = 30

        builder.Writeln("I'm a wonderful formatted cell.")

        builder.EndRow()
        builder.EndTable()

        dataDir = dataDir & Convert.ToString("DocumentBuilderSetTableCellFormatting_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderSetTableCellFormatting
        Console.WriteLine(Convert.ToString(vbLf & "Table cell formatting using DocumentBuilder set successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub SetTableRowFormatting(dataDir As String)
        ' ExStart:DocumentBuilderSetTableRowFormatting
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim table As Table = builder.StartTable()
        builder.InsertCell()

        ' Set the row formatting
        Dim rowFormat As RowFormat = builder.RowFormat
        rowFormat.Height = 100
        rowFormat.HeightRule = HeightRule.Exactly
        ' These formatting properties are set on the table and are applied to all rows in the table.
        table.LeftPadding = 30
        table.RightPadding = 30
        table.TopPadding = 30
        table.BottomPadding = 30

        builder.Writeln("I'm a wonderful formatted row.")

        builder.EndRow()
        builder.EndTable()

        dataDir = dataDir & Convert.ToString("DocumentBuilderSetTableRowFormatting_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderSetTableRowFormatting
        Console.WriteLine(Convert.ToString(vbLf & "Table row formatting using DocumentBuilder set successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub SetMultilevelListFormatting(dataDir As String)
        ' ExStart:DocumentBuilderSetMultilevelListFormatting
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.ListFormat.ApplyNumberDefault()

        builder.Writeln("Item 1")
        builder.Writeln("Item 2")

        builder.ListFormat.ListIndent()

        builder.Writeln("Item 2.1")
        builder.Writeln("Item 2.2")

        builder.ListFormat.ListIndent()

        builder.Writeln("Item 2.2.1")
        builder.Writeln("Item 2.2.2")

        builder.ListFormat.ListOutdent()

        builder.Writeln("Item 2.3")

        builder.ListFormat.ListOutdent()

        builder.Writeln("Item 3")

        builder.ListFormat.RemoveNumbers()
        dataDir = dataDir & Convert.ToString("DocumentBuilderSetMultilevelListFormatting_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderSetMultilevelListFormatting
        Console.WriteLine(Convert.ToString(vbLf & "Multilevel list formatting using DocumentBuilder set successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub SetPageSetupAndSectionFormatting(dataDir As String)
        ' ExStart:DocumentBuilderSetPageSetupAndSectionFormatting
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Set page properties
        builder.PageSetup.Orientation = Orientation.Landscape
        builder.PageSetup.LeftMargin = 50
        builder.PageSetup.PaperSize = PaperSize.Paper10x14

        dataDir = dataDir & Convert.ToString("DocumentBuilderSetPageSetupAndSectionFormatting_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderSetPageSetupAndSectionFormatting
        Console.WriteLine(Convert.ToString(vbLf & "Page setup and section formatting using DocumentBuilder set successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub ApplyParagraphStyle(dataDir As String)
        ' ExStart:DocumentBuilderApplyParagraphStyle
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Set paragraph style
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title

        builder.Write("Hello")
        dataDir = dataDir & Convert.ToString("DocumentBuilderApplyParagraphStyle_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderApplyParagraphStyle
        Console.WriteLine(Convert.ToString(vbLf & "Paragraph style using DocumentBuilder applied successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub ApplyBordersAndShadingToParagraph(dataDir As String)
        ' ExStart:DocumentBuilderApplyBordersAndShadingToParagraph
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Set paragraph borders
        Dim borders As BorderCollection = builder.ParagraphFormat.Borders
        borders.DistanceFromText = 20
        borders(BorderType.Left).LineStyle = LineStyle.[Double]
        borders(BorderType.Right).LineStyle = LineStyle.[Double]
        borders(BorderType.Top).LineStyle = LineStyle.[Double]
        borders(BorderType.Bottom).LineStyle = LineStyle.[Double]

        ' Set paragraph shading
        Dim shading As Shading = builder.ParagraphFormat.Shading
        shading.Texture = TextureIndex.TextureDiagonalCross
        shading.BackgroundPatternColor = System.Drawing.Color.LightCoral
        shading.ForegroundPatternColor = System.Drawing.Color.LightSalmon

        builder.Write("I'm a formatted paragraph with double border and nice shading.")
        dataDir = dataDir & Convert.ToString("DocumentBuilderApplyBordersAndShadingToParagraph_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderApplyBordersAndShadingToParagraph
        Console.WriteLine(Convert.ToString(vbLf & "Borders and shading using DocumentBuilder applied successfully to paragraph." & vbLf & "File saved at ") & dataDir)
    End Sub

End Class
