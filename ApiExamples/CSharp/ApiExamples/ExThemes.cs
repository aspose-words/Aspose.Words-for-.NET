// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Themes;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExThemes : ApiExampleBase
    {
        [Test]
        public void ThemeColors()
        {
            //ExStart
            //ExFor:Themes.ThemeColors
            //ExFor:Themes.ThemeColors.Accent1
            //ExFor:Themes.ThemeColors.Accent2
            //ExFor:Themes.ThemeColors.Accent3
            //ExFor:Themes.ThemeColors.Accent4
            //ExFor:Themes.ThemeColors.Accent5
            //ExFor:Themes.ThemeColors.Accent6
            //ExFor:Themes.ThemeColors.Dark1
            //ExFor:Themes.ThemeColors.Dark2
            //ExFor:Themes.ThemeColors.FollowedHyperlink
            //ExFor:Themes.ThemeColors.Hyperlink
            //ExFor:Themes.ThemeColors.Light1
            //ExFor:Themes.ThemeColors.Light2
            //ExSummary:Shows how to set custom theme colors.
            Document doc = new Document(MyDir + "ThemeColors.docx");

            // This collection of colors corresponds to the color palette from Microsoft Word which appears when changing shading or font color 
            ThemeColors colors = doc.Theme.Colors;

            colors.Dark1 = Color.MidnightBlue;
            colors.Light1 = Color.PaleGreen;
            colors.Dark2 = Color.Indigo;
            colors.Light2 = Color.Khaki;

            colors.Accent1 = Color.OrangeRed;
            colors.Accent2 = Color.LightSalmon;
            colors.Accent3 = Color.Yellow;
            colors.Accent4 = Color.Gold;
            colors.Accent5 = Color.BlueViolet;
            colors.Accent6 = Color.DarkViolet;

            // We can also set colors for hyperlinks like this
            colors.Hyperlink = Color.Black;
            colors.FollowedHyperlink = Color.Gray;

            doc.Save(ArtifactsDir + "Themes.ThemeColors.docx");
            //ExEnd
        }

        [Test]
        public void DocumentThemeProperties()
        {
            //ExStart
            //ExFor:Theme
            //ExFor:Theme.Colors
            //ExFor:Theme.MajorFonts
            //ExFor:Theme.MinorFonts
            //ExFor:Themes.ThemeFonts
            //ExFor:Themes.ThemeFonts.ComplexScript
            //ExFor:Themes.ThemeFonts.EastAsian
            //ExFor:Themes.ThemeFonts.Latin
            //ExSummary:Show how to change document theme options.
            Document doc = new Document();
            // Get document theme and do something useful
            Theme theme = doc.Theme;

            theme.Colors.Accent1 = Color.Black;
            theme.Colors.Dark1 = Color.Blue;
            theme.Colors.FollowedHyperlink = Color.White;
            theme.Colors.Hyperlink = Color.WhiteSmoke;
            theme.Colors.Light1 = Color.Empty; //There is default Color.Black

            theme.MajorFonts.ComplexScript = "Arial";
            theme.MajorFonts.EastAsian = String.Empty;
            theme.MajorFonts.Latin = "Times New Roman";

            theme.MinorFonts.ComplexScript = String.Empty;
            theme.MinorFonts.EastAsian = "Times New Roman";
            theme.MinorFonts.Latin = "Arial";
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual(Color.Black.ToArgb(), doc.Theme.Colors.Accent1.ToArgb());
            Assert.AreEqual(Color.Blue.ToArgb(), doc.Theme.Colors.Dark1.ToArgb());
            Assert.AreEqual(Color.White.ToArgb(), doc.Theme.Colors.FollowedHyperlink.ToArgb());
            Assert.AreEqual(Color.WhiteSmoke.ToArgb(), doc.Theme.Colors.Hyperlink.ToArgb());
            Assert.AreEqual(Color.Black.ToArgb(), doc.Theme.Colors.Light1.ToArgb());

            Assert.AreEqual("Arial", doc.Theme.MajorFonts.ComplexScript);
            Assert.AreEqual(String.Empty, doc.Theme.MajorFonts.EastAsian);
            Assert.AreEqual("Times New Roman", doc.Theme.MajorFonts.Latin);

            Assert.AreEqual(String.Empty, doc.Theme.MinorFonts.ComplexScript);
            Assert.AreEqual("Times New Roman", doc.Theme.MinorFonts.EastAsian);
            Assert.AreEqual("Arial", doc.Theme.MinorFonts.Latin);
        }

        [Test]
        public void DocTheme()
        {
            //ExStart
            //ExFor:Document.Theme
            //ExSummary:Shows what we can do with the Themes property of Document.
            Document doc = new Document();

            // When creating a blank document, Aspose Words creates a default theme object
            Theme theme = doc.Theme;

            // These color properties correspond to the 10 color columns that you see 
            // in the "Theme colors" section in the color selector menu when changing font or shading color
            // We can view and edit the leading color for each column, and the five different tints that
            // make up the rest of the column will be derived automatically from each leading color
            // Aspose Words sets the defaults to what they are in the Microsoft Word default theme
            Assert.AreEqual(Color.FromArgb(255, 255, 255, 255), theme.Colors.Light1);
            Assert.AreEqual(Color.FromArgb(255, 0, 0, 0), theme.Colors.Dark1);
            Assert.AreEqual(Color.FromArgb(255, 238, 236, 225), theme.Colors.Light2);
            Assert.AreEqual(Color.FromArgb(255, 31, 73, 125), theme.Colors.Dark2);
            Assert.AreEqual(Color.FromArgb(255, 79, 129, 189), theme.Colors.Accent1);
            Assert.AreEqual(Color.FromArgb(255, 192, 80, 77), theme.Colors.Accent2);
            Assert.AreEqual(Color.FromArgb(255, 155, 187, 89), theme.Colors.Accent3);
            Assert.AreEqual(Color.FromArgb(255, 128, 100, 162), theme.Colors.Accent4);
            Assert.AreEqual(Color.FromArgb(255, 75, 172, 198), theme.Colors.Accent5);
            Assert.AreEqual(Color.FromArgb(255, 247, 150, 70), theme.Colors.Accent6);

            // Hyperlink colors
            Assert.AreEqual(Color.FromArgb(255, 0, 0, 255), theme.Colors.Hyperlink);
            Assert.AreEqual(Color.FromArgb(255, 128, 0, 128), theme.Colors.FollowedHyperlink);

            // These appear at the very top of the font selector in the "Theme Fonts" section
            Assert.AreEqual("Cambria", theme.MajorFonts.Latin);
            Assert.AreEqual("Calibri", theme.MinorFonts.Latin);

            // Change some values to make a custom theme
            theme.MinorFonts.Latin = "Bodoni MT";
            theme.MajorFonts.Latin = "Tahoma";
            theme.Colors.Accent1 = Color.Cyan;
            theme.Colors.Accent2 = Color.Yellow;

            // Save the document to use our theme
            doc.Save(ArtifactsDir + "Document.Theme.docx");
            //ExEnd
        }
    }
}