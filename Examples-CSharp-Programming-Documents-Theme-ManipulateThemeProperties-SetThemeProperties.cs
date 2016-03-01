// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document doc = new Document(dataDir);
Theme theme = doc.Theme;
// Set Times New Roman font as Body theme font for Latin Character.
theme.MinorFonts.Latin = "Times New Roman";
// Set Color.Gold for theme color Hyperlink.
theme.Colors.Hyperlink = Color.Gold;            
