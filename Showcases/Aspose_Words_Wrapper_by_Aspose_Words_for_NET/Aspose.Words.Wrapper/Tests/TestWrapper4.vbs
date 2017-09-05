'
' Simple test with DocumentBuilder and Load/Save Options
'
Sub TestWrapper4
    Set WDocumentFactory = CreateObject("Aspose.Words.Wrapper.WDocumentFactory")
    SET WObjectFactory = CreateObject("Aspose.Words.Wrapper.WObjectFactory")

    Set LoadOptions = WObjectFactory.CreateObject("LoadOptions")
    LoadOptions.Password = "1234"
    Set WDoc = WDocumentFactory.OpenFromFileWithOptions("TestWrapper.docx", Options) ' this is user-defined method	

    Set WBuilder = CreateObject("Aspose.Words.Wrapper.WDocumentBuilder")
    Set WBuilder.Document = WDoc

    WBuilder.MoveToDocumentEnd
    WBuilder.WriteNewLine ' this is user-defined method
    WBuilder.WriteLine("Hi there again!") ' this is user-defined method
    WBuilder.Write("PAGE=")
    WBuilder.InsertFieldByFieldCode("PAGE") ' this is user-defined method

    Set SaveOptions = WObjectFactory.CreateObject("Saving.OoxmlSaveOptions")
    SaveOptions.Password = "4321"
    Call WDoc.SaveToFileWithOptions ("TestWrapper4.docx", SaveOptions)
End Sub

Call TestWrapper4