' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' When you call UpdatePageLayout the document is rendered in memory. Any warnings that occured during rendering
' are stored until the document save and then sent to the appropriate WarningCallback.
doc.UpdatePageLayout()

' Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
Dim callback As New HandleDocumentWarnings()

doc.WarningCallback = callback
dataDir = dataDir & Convert.ToString("Rendering.FontsNotificationUpdatePageLayout_out_.pdf")
' Even though the document was rendered previously, any save warnings are notified to the user during document save.
doc.Save(dataDir)
