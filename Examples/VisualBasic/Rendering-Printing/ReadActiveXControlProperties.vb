'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Drawing
Imports System.Drawing.Imaging

Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Rendering
Imports Aspose.Words.Saving
Imports Aspose.Words.Tables
Imports Aspose.Words.Drawing.Ole


Public Class ReadActiveXControlProperties
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

        ' Load the documents which store the shapes we want to render.           
        Dim doc As New Document(dataDir & Convert.ToString("ActiveXControl.docx"))

        Dim properties As String = ""
        ' Retrieve shapes from the document.         
        For Each shape As Shape In doc.GetChildNodes(NodeType.Shape, True)
            Dim oleControl As OleControl = shape.OleFormat.OleControl
            If oleControl.IsForms2OleControl Then
                Dim checkBox As Forms2OleControl = DirectCast(oleControl, Forms2OleControl)
                properties = (properties & Convert.ToString(vbLf & "Caption: ")) + checkBox.Caption
                properties = (properties & Convert.ToString(vbLf & "Value: ")) + checkBox.Value
                properties = (properties & Convert.ToString(vbLf & "Enabled: ")) + checkBox.Enabled.ToString
                properties = (properties & Convert.ToString(vbLf & "Type: ")) + checkBox.Type.ToString
                If checkBox.ChildNodes IsNot Nothing Then
                    properties = (properties & Convert.ToString(vbLf & "ChildNodes: ")) + checkBox.ChildNodes.ToString
                End If
                properties = properties & Convert.ToString(vbLf)
            End If
        Next
        properties = (properties & Convert.ToString(vbLf & "Total ActiveX Controls found: ")) + doc.GetChildNodes(NodeType.Shape, True).Count.ToString()
        Console.WriteLine(Convert.ToString(vbLf) & properties)
    End Sub
End Class
