Imports System
Imports System.Drawing
Imports System.Globalization
Imports Aspose.BarCode
Imports Aspose.Words.Fields
Imports Microsoft.VisualBasic


''' <summary>
''' Sample of custom barcode generator implementation (with underlying Aspose.BarCode module)
''' </summary>
Public Class CustomBarcodeGenerator
    Implements IBarcodeGenerator
    ''' <summary>
    ''' Converts barcode type from Word to Aspose.BarCode.
    ''' </summary>
    Private Shared Function ConvertBarcodeType(ByVal inputCode As String) As Symbology
        If inputCode Is Nothing Then
            Return CType(Integer.MinValue, Symbology)
        End If

        Dim type As String = inputCode.ToUpper()
        Dim outputCode As Symbology = CType(Integer.MinValue, Symbology)

        Select Case type
            Case "QR"
                outputCode = Symbology.QR
            Case "CODE128"
                outputCode = Symbology.Code128
            Case "CODE39"
                outputCode = Symbology.Code39Standard
            Case "EAN8"
                outputCode = Symbology.EAN8
            Case "EAN13"
                outputCode = Symbology.EAN13
            Case "UPCA"
                outputCode = Symbology.UPCA
            Case "UPCE"
                outputCode = Symbology.UPCE
            Case "ITF14"
                outputCode = Symbology.ITF14
            Case "CASE"
        End Select

        Return outputCode
    End Function

    ''' <summary>
    ''' Converts barcode image height from Word units to Aspose.BarCode units.
    ''' </summary>
    ''' <param name="heightInTwipsString"></param>
    ''' <returns></returns>
    Private Shared Function ConvertSymbolHeight(ByVal heightInTwipsString As String) As Single
        ' Input value is in 1/1440 inches (twips)
        Dim heightInTwips As Integer = TryParseInt(heightInTwipsString)
        If heightInTwips = Integer.MinValue Then
            Throw New Exception("Error! Incorrect height - " & heightInTwipsString & ".")
        End If

        ' Convert to mm
        Return CSng(heightInTwips * 25.4 / 1440)
    End Function

    ''' <summary>
    ''' Converts barcode image color from Word to Aspose.BarCode.
    ''' </summary>
    ''' <param name="inputColor"></param>
    ''' <returns></returns>
    Private Shared Function ConvertColor(ByVal inputColor As String) As Color
        ' Input should be from "0x000000" to "0xFFFFFF"
        Dim color As Integer = TryParseHex(inputColor.Replace("0x", ""))
        If color = Integer.MinValue Then
            Throw New Exception("Error! Incorrect color - " & inputColor & ".")
        End If

        Return Drawing.Color.FromArgb(color >> 16, (color And &HFF00) >> 8, color And &HFF)
    End Function

    ''' <summary>
    ''' Converts bar code scaling factor from percents to float.
    ''' </summary>
    ''' <param name="scalingFactor"></param>
    ''' <returns></returns>
    Private Shared Function ConvertScalingFactor(ByVal scalingFactor As String) As Single
        Dim isParsed As Boolean = False
        Dim percents As Integer = TryParseInt(scalingFactor)

        If percents <> Integer.MinValue Then
            If percents >= 10 AndAlso percents <= 10000 Then
                isParsed = True
            End If
        End If

        If (Not isParsed) Then
            Throw New Exception("Error! Incorrect scaling factor - " & scalingFactor & ".")
        End If

        Return percents / 100.0F
    End Function
    ''' <summary>
    ''' Implementation of the GetOldBarcodeImage() method for IBarCodeGenerator interface.
    ''' </summary>
    ''' <param name="parameters"></param>
    ''' <returns></returns>
    Public Function IBarcodeGenerator_GetOldBarcodeImage(ByVal parameters As BarcodeParameters) As Image Implements IBarcodeGenerator.GetOldBarcodeImage
        If parameters.PostalAddress Is Nothing Then
            Return Nothing
        End If

        Dim builder As New BarCodeBuilder()

        ' Hardcode type for old-fashioned Barcode
        builder.SymbologyType = Symbology.Postnet
        builder.CodeText = parameters.PostalAddress

        Return builder.BarCodeImage
    End Function
    ''' <summary>
    ''' Implementation of the GetBarcodeImage() method for IBarCodeGenerator interface.
    ''' </summary>
    ''' <param name="parameters"></param>
    ''' <returns></returns>
    Public Function IBarcodeGenerator_GetBarcodeImage(ByVal parameters As BarcodeParameters) As Image Implements IBarcodeGenerator.GetBarcodeImage
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

        Const scale As Single = 0.4F ' Empiric scaling factor for converting Word barcode to Aspose.BarCode
        Dim xdim As Single = 1.0F

        If builder.SymbologyType = Symbology.QR Then
            builder.AutoSize = False
            builder.ImageWidth *= scale
            builder.ImageHeight = builder.ImageWidth
            xdim = builder.ImageHeight / 25
            builder.yDimension = xdim
            builder.xDimension = xdim
        End If

        If parameters.ScalingFactor IsNot Nothing Then
            Dim scalingFactor As Single = ConvertScalingFactor(parameters.ScalingFactor)
            builder.ImageHeight *= scalingFactor
            If builder.SymbologyType = Symbology.QR Then
                builder.ImageWidth = builder.ImageHeight
                builder.yDimension = xdim * scalingFactor
                builder.xDimension = builder.yDimension
            End If

            builder.AutoSize = False
        End If

        Return builder.BarCodeImage
    End Function
    ''' <summary>
    ''' Parses an integer using the invariant culture. Returns Int.MinValue if cannot parse.
    ''' 
    ''' Allows leading sign.
    ''' Allows leading and trailing spaces.
    ''' </summary>
    Public Shared Function TryParseInt(ByVal s As String) As Integer
        Dim temp As Double
        Return If((Double.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, temp)), CastDoubleToInt(temp), Integer.MinValue)
    End Function

    ''' <summary>
    ''' Casts a double to int32 in a way that uint32 are "correctly" casted too (they become negative numbers).
    ''' </summary>
    Public Shared Function CastDoubleToInt(ByVal value As Double) As Integer
        Dim temp As Long = CLng(Fix(value))
        Return CInt(Fix(temp))
    End Function

    ''' <summary>
    ''' Try parses a hex string into an integer value.
    ''' on error return int.MinValue
    ''' </summary>
    Public Shared Function TryParseHex(ByVal s As String) As Integer
        Dim result As Integer
        Return If(Integer.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, result), result, Integer.MinValue)
    End Function
End Class