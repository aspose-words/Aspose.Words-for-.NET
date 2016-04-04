// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
// To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false.
PdfSaveOptions options = new PdfSaveOptions();
options.EmbedStandardWindowsFonts = false;

dataDir = dataDir + "Rendering.DisableEmbedWindowsFonts_out_.pdf";
// The output PDF will be saved without embedding standard windows fonts.
doc.Save(dataDir);
