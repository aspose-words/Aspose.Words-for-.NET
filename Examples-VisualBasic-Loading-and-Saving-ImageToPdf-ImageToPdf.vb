' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

ConvertImageToPdf(dataDir & Convert.ToString("Test.jpg"), dataDir & Convert.ToString("TestJpg_out_.pdf"))
ConvertImageToPdf(dataDir & Convert.ToString("Test.png"), dataDir & Convert.ToString("TestPng_out_.pdf"))
ConvertImageToPdf(dataDir & Convert.ToString("Test.wmf"), dataDir & Convert.ToString("TestWmf_out_.pdf"))
ConvertImageToPdf(dataDir & Convert.ToString("Test.tiff"), dataDir & Convert.ToString("TestTiff_out_.pdf"))
ConvertImageToPdf(dataDir & Convert.ToString("Test.gif"), dataDir & Convert.ToString("TestGif_out_.pdf"))
