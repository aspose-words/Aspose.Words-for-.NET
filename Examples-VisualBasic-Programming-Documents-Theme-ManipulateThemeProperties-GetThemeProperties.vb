' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document(dataDir)
Dim theme As Theme = doc.Theme
' Major (Headings) font for Latin characters.
Console.WriteLine(theme.MajorFonts.Latin)
' Minor (Body) font for EastAsian characters.
Console.WriteLine(theme.MinorFonts.EastAsian)
' Color for theme color Accent 1.
Console.WriteLine(theme.Colors.Accent1)
