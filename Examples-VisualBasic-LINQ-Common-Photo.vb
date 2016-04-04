' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_LINQ()

' Load the photo and read all bytes.
Dim imgdata As Byte() = System.IO.File.ReadAllBytes(dataDir & Convert.ToString("photo.png"))
Return imgdata
