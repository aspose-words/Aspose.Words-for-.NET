'
' Simple test with DOM
'
Sub TestWrapper2
    Set WDocumentFactory = CreateObject("Aspose.Words.Wrapper.WDocumentFactory")
    Set WDoc = WDocumentFactory.OpenFromFile("TestWrapper.docx") ' this is user-defined method

    Set WPar1 = WDoc.CreateNode("WParagraph") ' Create wrapper object

    Set Run1 = WDoc.CreateNode("Run") ' Create ordinary object
    Run1.Text = "Page "

    Set Run2 = WDoc.CreateNode("Run")
    Run2.Text = " from "

    WPar1.AppendChild(Run1)
    WPar1.AppendFieldByCode("PAGE") ' this is user-defined method
    WPar1.AppendChild(Run2)
    WPar1.AppendFieldByCode("NUMPAGES") ' this is user-defined method

    WDoc.LastSection.Body.AppendChild(WPar1)

    WDoc.SaveToFile("TestWrapper2.docx") ' this is user-defined method
End Sub

Call TestWrapper2