// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using Aspose.Words;
using Aspose.Words.Themes;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExThemes : ApiExampleBase
    {
        [Test]
        public void CustomColorsAndFonts()
        {
            //ExStart
            //ExFor:Document.Theme
            //ExFor:Theme
            //ExFor:Theme.Colors
            //ExFor:Theme.MajorFonts
            //ExFor:Theme.MinorFonts
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
            //ExFor:Themes.ThemeFonts
            //ExFor:Themes.ThemeFonts.ComplexScript
            //ExFor:Themes.ThemeFonts.EastAsian
            //ExFor:Themes.ThemeFonts.Latin
            //ExSummary:Shows how to set custom theme colors and fonts.
            Document doc = new Document(MyDir + "ThemeColors.docx");

            // This object gives us access to the document theme, which is a source of default fonts and colors
            Theme theme = doc.Theme;

            // These fonts will be inherited by some styles like "Heading 1" and "Subtitle"
            theme.MajorFonts.Latin = "Courier New";
            theme.MinorFonts.Latin = "Agency FB";

            Assert.AreEqual(string.Empty, theme.MajorFonts.ComplexScript);
            Assert.AreEqual(string.Empty, theme.MajorFonts.EastAsian);
            Assert.AreEqual(string.Empty, theme.MinorFonts.ComplexScript);
            Assert.AreEqual(string.Empty, theme.MinorFonts.EastAsian);

            // This collection of colors corresponds to the color palette from Microsoft Word which appears when changing shading or font color 
            ThemeColors colors = theme.Colors;

            // We will set the color of each color palette column going from left to right like this
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

            doc.Save(ArtifactsDir + "Themes.CustomColorsAndFonts.docx");
            //ExEnd
            doc = new Document(ArtifactsDir + "Themes.CustomColorsAndFonts.docx");

            Assert.AreEqual(Color.OrangeRed.ToArgb(), doc.Theme.Colors.Accent1.ToArgb());
            Assert.AreEqual(Color.MidnightBlue.ToArgb(), doc.Theme.Colors.Dark1.ToArgb());
            Assert.AreEqual(Color.Gray.ToArgb(), doc.Theme.Colors.FollowedHyperlink.ToArgb());
            Assert.AreEqual(Color.Black.ToArgb(), doc.Theme.Colors.Hyperlink.ToArgb());
            Assert.AreEqual(Color.PaleGreen.ToArgb(), doc.Theme.Colors.Light1.ToArgb());

            Assert.AreEqual(string.Empty, doc.Theme.MajorFonts.ComplexScript);
            Assert.AreEqual(string.Empty, doc.Theme.MajorFonts.EastAsian);
            Assert.AreEqual("Courier New", doc.Theme.MajorFonts.Latin);

            Assert.AreEqual(string.Empty, doc.Theme.MinorFonts.ComplexScript);
            Assert.AreEqual(string.Empty, doc.Theme.MinorFonts.EastAsian);
            Assert.AreEqual("Agency FB", doc.Theme.MinorFonts.Latin);
        }
    }
}