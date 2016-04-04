' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' Check the file format and move the file to the appropriate folder.
Dim info As FileFormatInfo = FileFormatUtil.DetectFileFormat(fileName)

' Display the document type.
Select Case info.LoadFormat
    Case LoadFormat.Doc
        Console.WriteLine(Constants.vbTab & "Microsoft Word 97-2003 document.")
    Case LoadFormat.Dot
        Console.WriteLine(Constants.vbTab & "Microsoft Word 97-2003 template.")
    Case LoadFormat.Docx
        Console.WriteLine(Constants.vbTab & "Office Open XML WordprocessingML Macro-Free Document.")
    Case LoadFormat.Docm
        Console.WriteLine(Constants.vbTab & "Office Open XML WordprocessingML Macro-Enabled Document.")
    Case LoadFormat.Dotx
        Console.WriteLine(Constants.vbTab & "Office Open XML WordprocessingML Macro-Free Template.")
    Case LoadFormat.Dotm
        Console.WriteLine(Constants.vbTab & "Office Open XML WordprocessingML Macro-Enabled Template.")
    Case LoadFormat.FlatOpc
        Console.WriteLine(Constants.vbTab & "Flat OPC document.")
    Case LoadFormat.Rtf
        Console.WriteLine(Constants.vbTab & "RTF format.")
    Case LoadFormat.WordML
        Console.WriteLine(Constants.vbTab & "Microsoft Word 2003 WordprocessingML format.")
    Case LoadFormat.Html
        Console.WriteLine(Constants.vbTab & "HTML format.")
    Case LoadFormat.Mhtml
        Console.WriteLine(Constants.vbTab & "MHTML (Web archive) format.")
    Case LoadFormat.Odt
        Console.WriteLine(Constants.vbTab & "OpenDocument Text.")
    Case LoadFormat.Ott
        Console.WriteLine(Constants.vbTab & "OpenDocument Text Template.")
    Case LoadFormat.DocPreWord97
        Console.WriteLine(Constants.vbTab & "MS Word 6 or Word 95 format.")
    Case Else
        Console.WriteLine(Constants.vbTab & "Unknown format.")
End Select
