Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Themes
Imports System.Drawing
Public Class ManipulateThemeProperties
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTheme()
        dataDir = dataDir & Convert.ToString("Document.doc")
        GetThemeProperties(dataDir)
        SetThemeProperties(dataDir)
    End Sub
    ''' <summary>
    '''  Shows how to get theme properties.
    ''' </summary>             
    Private Shared Sub GetThemeProperties(dataDir As String)
        ' ExStart:GetThemeProperties
        Dim doc As New Document(dataDir)
        Dim theme As Theme = doc.Theme
        ' Major (Headings) font for Latin characters.
        Console.WriteLine(theme.MajorFonts.Latin)
        ' Minor (Body) font for EastAsian characters.
        Console.WriteLine(theme.MinorFonts.EastAsian)
        ' Color for theme color Accent 1.
        Console.WriteLine(theme.Colors.Accent1)
        ' ExEnd:GetThemeProperties 
    End Sub
    ''' <summary>
    '''  Shows how to set theme properties.
    ''' </summary>             
    Private Shared Sub SetThemeProperties(dataDir As String)
        ' ExStart:SetThemeProperties
        Dim doc As New Document(dataDir)
        Dim theme As Theme = doc.Theme
        ' Set Times New Roman font as Body theme font for Latin Character.
        theme.MinorFonts.Latin = "Times New Roman"
        ' Set Color.Gold for theme color Hyperlink.
        theme.Colors.Hyperlink = Color.Gold
        ' ExEnd:SetThemeProperties 
    End Sub
End Class
