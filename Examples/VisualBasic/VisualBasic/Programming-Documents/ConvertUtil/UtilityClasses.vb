Imports Aspose.Words
Public Class UtilityClasses
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithHyperlink()
        ConvertBetweenMeasurementUnits()
        UseControlCharacters()
    End Sub

    Private Shared Sub ConvertBetweenMeasurementUnits()
        ' ExStart:ConvertBetweenMeasurementUnits
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim pageSetup As PageSetup = builder.PageSetup
        pageSetup.TopMargin = ConvertUtil.InchToPoint(1.0)
        pageSetup.BottomMargin = ConvertUtil.InchToPoint(1.0)
        pageSetup.LeftMargin = ConvertUtil.InchToPoint(1.5)
        pageSetup.RightMargin = ConvertUtil.InchToPoint(1.5)
        pageSetup.HeaderDistance = ConvertUtil.InchToPoint(0.2)
        pageSetup.FooterDistance = ConvertUtil.InchToPoint(0.2)
        ' ExEnd:ConvertBetweenMeasurementUnits
        Console.WriteLine(vbLf & "Page properties specified in inches.")

    End Sub
    Private Shared Sub UseControlCharacters()
        ' ExStart:UseControlCharacters
        Dim text As String = "test" & vbCr
        ' Replace "\r" control character with "\r\n"
        text = text.Replace(ControlChar.Cr, ControlChar.CrLf)
        ' ExEnd:UseControlCharacters
        Console.WriteLine(vbLf & "Control characters used successfully.")

    End Sub
End Class
