'
' Simple test with DocumentBuilder
'
Sub TestWrapper1
    Set oLicense = CreateObject("Aspose.Words.License")
    'oLicense.SetLicense("Put.License.File.Name.Here")
    isLicenseSet = oLicense.IsLicensed
    MsgBox("IsLicensed=" & isLicenseSet)

    Set WDocumentFactory = CreateObject("Aspose.Words.Wrapper.WDocumentFactory")
    Set WDoc = WDocumentFactory.OpenFromFile("TestWrapper.docx") ' this is user-defined method

    Set WBuilder = CreateObject("Aspose.Words.Wrapper.WDocumentBuilder")
    Set WBuilder.Document = WDoc 

    WBuilder.MoveToDocumentEnd
    WBuilder.WriteNewLine ' this is user-defined method
    WBuilder.WriteLine("Hi there again!") ' this is user-defined method
    WBuilder.Write("PAGE=")
    WBuilder.InsertFieldByFieldCode("PAGE") ' this is user-defined method

    WDoc.SaveToFile("TestWrapper1.docx") ' this is user-defined method
End Sub

Call TestWrapper1