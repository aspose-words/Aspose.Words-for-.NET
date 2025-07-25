﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
            //ExFor:ThemeColors
            //ExFor:ThemeColors.Accent1
            //ExFor:ThemeColors.Accent2
            //ExFor:ThemeColors.Accent3
            //ExFor:ThemeColors.Accent4
            //ExFor:ThemeColors.Accent5
            //ExFor:ThemeColors.Accent6
            //ExFor:ThemeColors.Dark1
            //ExFor:ThemeColors.Dark2
            //ExFor:ThemeColors.FollowedHyperlink
            //ExFor:ThemeColors.Hyperlink
            //ExFor:ThemeColors.Light1
            //ExFor:ThemeColors.Light2
            //ExFor:ThemeFonts
            //ExFor:ThemeFonts.ComplexScript
            //ExFor:ThemeFonts.EastAsian
            //ExFor:ThemeFonts.Latin
            //ExSummary:Shows how to set custom colors and fonts for themes.
            Document doc = new Document(MyDir + "Theme colors.docx");

            // The "Theme" object gives us access to the document theme, a source of default fonts and colors.
            Theme theme = doc.Theme;

            // Some styles, such as "Heading 1" and "Subtitle", will inherit these fonts.
            theme.MajorFonts.Latin = "Courier New";
            theme.MinorFonts.Latin = "Agency FB";

            // Other languages may also have their custom fonts in this theme.
            Assert.That(theme.MajorFonts.ComplexScript, Is.EqualTo(string.Empty));
            Assert.That(theme.MajorFonts.EastAsian, Is.EqualTo(string.Empty));
            Assert.That(theme.MinorFonts.ComplexScript, Is.EqualTo(string.Empty));
            Assert.That(theme.MinorFonts.EastAsian, Is.EqualTo(string.Empty));

            // The "Colors" property contains the color palette from Microsoft Word,
            // which appears when changing shading or font color.
            // Apply custom colors to the color palette so we have easy access to them in Microsoft Word
            // when we, for example, change the font color via "Home" -> "Font" -> "Font Color",
            // or insert a shape, and then set a color for it via "Shape Format" -> "Shape Styles".
            ThemeColors colors = theme.Colors;
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

            // Apply custom colors to hyperlinks in their clicked and un-clicked states.
            colors.Hyperlink = Color.Black;
            colors.FollowedHyperlink = Color.Gray;

            doc.Save(ArtifactsDir + "Themes.CustomColorsAndFonts.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Themes.CustomColorsAndFonts.docx");

            Assert.That(doc.Theme.Colors.Accent1.ToArgb(), Is.EqualTo(Color.OrangeRed.ToArgb()));
            Assert.That(doc.Theme.Colors.Dark1.ToArgb(), Is.EqualTo(Color.MidnightBlue.ToArgb()));
            Assert.That(doc.Theme.Colors.FollowedHyperlink.ToArgb(), Is.EqualTo(Color.Gray.ToArgb()));
            Assert.That(doc.Theme.Colors.Hyperlink.ToArgb(), Is.EqualTo(Color.Black.ToArgb()));
            Assert.That(doc.Theme.Colors.Light1.ToArgb(), Is.EqualTo(Color.PaleGreen.ToArgb()));

            Assert.That(doc.Theme.MajorFonts.ComplexScript, Is.EqualTo(string.Empty));
            Assert.That(doc.Theme.MajorFonts.EastAsian, Is.EqualTo(string.Empty));
            Assert.That(doc.Theme.MajorFonts.Latin, Is.EqualTo("Courier New"));

            Assert.That(doc.Theme.MinorFonts.ComplexScript, Is.EqualTo(string.Empty));
            Assert.That(doc.Theme.MinorFonts.EastAsian, Is.EqualTo(string.Empty));
            Assert.That(doc.Theme.MinorFonts.Latin, Is.EqualTo("Agency FB"));
        }
    }
}