

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Public Class CheckFormat
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()
        Dim supportedDir As String = dataDir & "OutSupported"
        Dim unknownDir As String = dataDir & "OutUnknown"
        Dim encryptedDir As String = dataDir & "OutEncrypted"
        Dim pre97Dir As String = dataDir & "OutPre97"


        ' Create the directories if they do not already exist
        If Directory.Exists(supportedDir) = False Then
            Directory.CreateDirectory(supportedDir)
        End If
        If Directory.Exists(unknownDir) = False Then
            Directory.CreateDirectory(unknownDir)
        End If
        If Directory.Exists(encryptedDir) = False Then
            Directory.CreateDirectory(encryptedDir)
        End If
        If Directory.Exists(pre97Dir) = False Then
            Directory.CreateDirectory(pre97Dir)
        End If

        Dim fileList() As String = Directory.GetFiles(dataDir)
        
        For Each fileName As String In fileList
            ' Extract and display the file name without the path.
            Dim nameOnly As String = Path.GetFileName(fileName)
            Console.Write(nameOnly)

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

            ' Now copy the document into the appropriate folder.
            If info.IsEncrypted Then
                Console.WriteLine(Constants.vbTab & "An encrypted document.")
                File.Copy(fileName, Path.Combine(encryptedDir, nameOnly), True)
            Else
                Select Case info.LoadFormat
                    Case LoadFormat.DocPreWord97
                        File.Copy(fileName, Path.Combine(pre97Dir, nameOnly), True)
                    Case LoadFormat.Unknown
                        File.Copy(fileName, Path.Combine(unknownDir, nameOnly), True)
                    Case Else
                        File.Copy(fileName, Path.Combine(supportedDir, nameOnly), True)
                End Select
            End If
        Next fileName

        Console.WriteLine(vbNewLine + "Checked the format of all documents successfully.")
    End Sub
End Class
