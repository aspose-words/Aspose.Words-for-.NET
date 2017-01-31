Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Imports System.Drawing
Imports Aspose.BarCode
Class GenerateACustomBarCodeImage
        Public Shared Sub Run()
            ' ExStart:GenerateACustomBarCodeImage
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
            Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))

            ' Set custom barcode generator
            doc.FieldOptions.BarcodeGenerator = New CustomBarcodeGenerator()
            doc.Save(dataDir & Convert.ToString("GenerateACustomBarCodeImage_out.pdf"))
            ' ExEnd:GenerateACustomBarCodeImage
        End Sub
    End Class

    ' ExStart:GenerateACustomBarCodeImage_IBarcodeGenerator
    Public Class CustomBarcodeGenerator
        Implements IBarcodeGenerator
        ''' <summary>
        ''' Converts barcode type from Word to Aspose.BarCode.
        ''' </summary>
        ''' <param name="inputCode"></param>
        ''' <returns></returns>
        Private Shared Function ConvertBarcodeType(inputCode As String) As Symbology
            If inputCode Is Nothing Then
                Return Integer.MinValue
            End If

            Dim type As String = inputCode.ToUpper()

            Select Case type
                Case "QR"
                    Return Symbology.QR
                Case "CODE128"
                    Return Symbology.Code128
                Case "CODE39"
                    Return Symbology.Code39Standard
                Case "EAN8"
                    Return Symbology.EAN8
                Case "EAN13"
                    Return Symbology.EAN13
                Case "UPCA"
                    Return Symbology.UPCA
                Case "UPCE"
                    Return Symbology.UPCE
                Case "ITF14"
                    Return Symbology.ITF14
                Case "CASE"
                    Exit Select
            End Select

            Return Integer.MinValue
        End Function

        ''' <summary>
        ''' Converts barcode image height from Word units to Aspose.BarCode units.
        ''' </summary>
        ''' <param name="heightInTwipsString"></param>
        ''' <returns></returns>
        Private Shared Function ConvertSymbolHeight(heightInTwipsString As String) As Single
            ' Input value is in 1/1440 inches (twips)
            Dim heightInTwips As Integer = Integer.MinValue
            Integer.TryParse(heightInTwipsString, heightInTwips)

            If heightInTwips = Integer.MinValue Then
                Throw New Exception((Convert.ToString("Error! Incorrect height - ") & heightInTwipsString) + ".")
            End If

            ' Convert to mm
            Return CSng(heightInTwips * 25.4 / 1440)
        End Function

        ''' <summary>
        ''' Converts barcode image color from Word to Aspose.BarCode.
        ''' </summary>
        ''' <param name="inputColor"></param>
        ''' <returns></returns>
        Private Shared Function ConvertColor(inputColor As String) As Color
            ' Input should be from "0x000000" to "0xFFFFFF"
            Dim color__1 As Integer = Integer.MinValue
            Integer.TryParse(inputColor.Replace("0x", ""), color__1)

            If color__1 = Integer.MinValue Then
                Throw New Exception((Convert.ToString("Error! Incorrect color - ") & inputColor) + ".")
            End If

            Return Color.FromArgb(color__1 >> 16, (color__1 And &HFF00) >> 8, color__1 And &HFF)

            ' Backword conversion -
            'return string.Format("0x{0,6:X6}", mControl.ForeColor.ToArgb() & 0xFFFFFF);
        End Function

        ''' <summary>
        ''' Converts bar code scaling factor from percents to float.
        ''' </summary>
        ''' <param name="scalingFactor"></param>
        ''' <returns></returns>
        Private Shared Function ConvertScalingFactor(scalingFactor As String) As Single
            Dim isParsed As Boolean = False
            Dim percents As Integer = Integer.MinValue
            Integer.TryParse(scalingFactor, percents)

            If percents <> Integer.MinValue Then
                If percents >= 10 AndAlso percents <= 10000 Then
                    isParsed = True
                End If
            End If

            If Not isParsed Then
                Throw New Exception((Convert.ToString("Error! Incorrect scaling factor - ") & scalingFactor) + ".")
            End If

            Return percents / 100.0F
        End Function

        ''' <summary>
        ''' Implementation of the GetBarCodeImage() method for IBarCodeGenerator interface.
        ''' </summary>
        ''' <param name="parameters"></param>
        ''' <returns></returns>
        Public Function GetBarcodeImage(parameters As BarcodeParameters) As Image
            If parameters.BarcodeType Is Nothing OrElse parameters.BarcodeValue Is Nothing Then
                Return Nothing
            End If

            Dim builder As New BarCodeBuilder()

            builder.SymbologyType = ConvertBarcodeType(parameters.BarcodeType)
            If builder.SymbologyType = CType(Integer.MinValue, Symbology) Then
                Return Nothing
            End If

            builder.CodeText = parameters.BarcodeValue

            If builder.SymbologyType = Symbology.QR Then
                builder.Display2DText = parameters.BarcodeValue
            End If

            If parameters.ForegroundColor IsNot Nothing Then
                builder.ForeColor = ConvertColor(parameters.ForegroundColor)
            End If

            If parameters.BackgroundColor IsNot Nothing Then
                builder.BackColor = ConvertColor(parameters.BackgroundColor)
            End If

            If parameters.SymbolHeight IsNot Nothing Then
                builder.ImageHeight = ConvertSymbolHeight(parameters.SymbolHeight)
                builder.AutoSize = False
            End If

            builder.CodeLocation = CodeLocation.None

            If parameters.DisplayText Then
                builder.CodeLocation = CodeLocation.Below
            End If

            builder.CaptionAbove.Text = ""

            Const scale As Single = 0.4F
            ' Empiric scaling factor for converting Word barcode to Aspose.BarCode
            Dim xdim As Single = 1.0F

            If builder.SymbologyType = Symbology.QR Then
                builder.AutoSize = False
                builder.ImageWidth *= scale
                builder.ImageHeight = builder.ImageWidth
                xdim = builder.ImageHeight / 25
                builder.xDimension = InlineAssignHelper(builder.yDimension, xdim)
            End If

            If parameters.ScalingFactor IsNot Nothing Then
                Dim scalingFactor As Single = ConvertScalingFactor(parameters.ScalingFactor)
                builder.ImageHeight *= scalingFactor
                If builder.SymbologyType = Symbology.QR Then
                    builder.ImageWidth = builder.ImageHeight
                    builder.xDimension = InlineAssignHelper(builder.yDimension, xdim * scalingFactor)
                End If

                builder.AutoSize = False
            End If
            Return builder.BarCodeImage
        End Function

        Private Function IBarcodeGenerator_GetBarcodeImage(parameters As BarcodeParameters) As Image Implements IBarcodeGenerator.GetBarcodeImage
            Throw New NotImplementedException()
        End Function

        Public Function GetOldBarcodeImage(parameters As BarcodeParameters) As Image
            Throw New NotImplementedException()
        End Function
        Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function

        Private Function IBarcodeGenerator_GetOldBarcodeImage(parameters As BarcodeParameters) As Image Implements IBarcodeGenerator.GetOldBarcodeImage
            Throw New NotImplementedException()
        End Function
    End Class
' ExEnd:GenerateACustomBarCodeImage_IBarcodeGenerator