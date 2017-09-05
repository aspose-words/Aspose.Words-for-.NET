'
' More complex test with DOM
'
Sub TestWrapper3
    Set WDocumentFactory = CreateObject("Aspose.Words.Wrapper.WDocumentFactory")
    Set WDoc = WDocumentFactory.OpenFromFile("TestWrapper.docx") ' this is user-defined method

    Set WHeader = WDoc.CreateNode("WHeader") ' Create object associated with WDoc
    WDoc.GetSection(0).AppendChild(WHeader) ' GetSection(N) is user-defined method

    Set WTab1 = WDoc.CreateNode("Tables.WTable") ' Create more objects associated with WDoc
    Set WRow1 = WDoc.CreateNode("Tables.WRow")

    WTab1.AppendChild (WRow1)
    Set WCell1 = WDoc.CreateNode("Tables.WCell")

    WRow1.AppendChild(WCell1)
    WHeader.AppendChild(WTab1)
    Set WPar1 = WDoc.CreateNode("WParagraph")
    WCell1.AppendChild(WPar1)

    Set WRun1 = WDoc.CreateNode("WRun")
    WPar1.AppendChild(WRun1)
    Set WRun1 = WPar1.GetRun(0)
    WRun1.Text = "This is header of the document. Created at "
    WPar1.AppendFieldByCode("CREATEDATE") ' this is user-defined method

    Set WFooter = WDoc.CreateNode("WFooter")
    WDoc.GetSection(0).AppendChild(WFooter)

    Set WPar2 = WDoc.CreateNode("WParagraph")
    WFooter.AppendChild(WPar2)

    Set WRun2 = WDoc.CreateNode("WRun")
    WRun2.Text = "Page "

    WPar2.AppendChild(WRun2)
    WPar2.AppendFieldByCode("PAGE") ' this is user-defined method
    Set WRun2 = WDoc.CreateNode("WRun")
    WRun2.Text = " from "
    WPar2.AppendChild(WRun2)
    WPar2.AppendFieldByCode("NUMPAGES") ' this is user-defined method

    WDoc.SaveToFile("TestWrapper3.docx") ' this is user-defined method
End Sub

Call TestWrapper3