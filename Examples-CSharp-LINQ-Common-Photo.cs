// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_LINQ();

// Load the photo and read all bytes.
byte[] imgdata = System.IO.File.ReadAllBytes(dataDir + "photo.png");
return imgdata;
