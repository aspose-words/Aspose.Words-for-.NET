// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExCompatibilityOptions : ApiExampleBase
    {
        //ExStart
        //ExFor:Compatibility
        //ExFor:CompatibilityOptions
        //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
        //ExFor:Document.CompatibilityOptions
        //ExFor:MsWordVersion
        //ExFor:CompatibilityOptions.AdjustLineHeightInTable
        //ExFor:CompatibilityOptions.AlignTablesRowByRow
        //ExFor:CompatibilityOptions.AllowSpaceOfSameStyleInTable
        //ExFor:CompatibilityOptions.ApplyBreakingRules
        //ExFor:CompatibilityOptions.AutofitToFirstFixedWidthCell
        //ExFor:CompatibilityOptions.AutoSpaceLikeWord95
        //ExFor:CompatibilityOptions.BalanceSingleByteDoubleByteWidth
        //ExFor:CompatibilityOptions.CachedColBalance
        //ExFor:CompatibilityOptions.ConvMailMergeEsc
        //ExFor:CompatibilityOptions.DisableOpenTypeFontFormattingFeatures
        //ExFor:CompatibilityOptions.DisplayHangulFixedWidth
        //ExFor:CompatibilityOptions.DoNotAutofitConstrainedTables
        //ExFor:CompatibilityOptions.DoNotBreakConstrainedForcedTable
        //ExFor:CompatibilityOptions.DoNotBreakWrappedTables
        //ExFor:CompatibilityOptions.DoNotExpandShiftReturn
        //ExFor:CompatibilityOptions.DoNotLeaveBackslashAlone
        //ExFor:CompatibilityOptions.DoNotSnapToGridInCell
        //ExFor:CompatibilityOptions.DoNotSuppressIndentation
        //ExFor:CompatibilityOptions.DoNotSuppressParagraphBorders
        //ExFor:CompatibilityOptions.DoNotUseEastAsianBreakRules
        //ExFor:CompatibilityOptions.DoNotUseHTMLParagraphAutoSpacing
        //ExFor:CompatibilityOptions.DoNotUseIndentAsNumberingTabStop
        //ExFor:CompatibilityOptions.DoNotVertAlignCellWithSp
        //ExFor:CompatibilityOptions.DoNotVertAlignInTxbx
        //ExFor:CompatibilityOptions.DoNotWrapTextWithPunct
        //ExFor:CompatibilityOptions.FootnoteLayoutLikeWW8
        //ExFor:CompatibilityOptions.ForgetLastTabAlignment
        //ExFor:CompatibilityOptions.GrowAutofit
        //ExFor:CompatibilityOptions.LayoutRawTableWidth
        //ExFor:CompatibilityOptions.LayoutTableRowsApart
        //ExFor:CompatibilityOptions.LineWrapLikeWord6
        //ExFor:CompatibilityOptions.MWSmallCaps
        //ExFor:CompatibilityOptions.NoColumnBalance
        //ExFor:CompatibilityOptions.NoExtraLineSpacing
        //ExFor:CompatibilityOptions.NoLeading
        //ExFor:CompatibilityOptions.NoSpaceRaiseLower
        //ExFor:CompatibilityOptions.NoTabHangInd
        //ExFor:CompatibilityOptions.OverrideTableStyleFontSizeAndJustification
        //ExFor:CompatibilityOptions.PrintBodyTextBeforeHeader
        //ExFor:CompatibilityOptions.PrintColBlack
        //ExFor:CompatibilityOptions.SelectFldWithFirstOrLastChar
        //ExFor:CompatibilityOptions.ShapeLayoutLikeWW8
        //ExFor:CompatibilityOptions.ShowBreaksInFrames
        //ExFor:CompatibilityOptions.SpaceForUL
        //ExFor:CompatibilityOptions.SpacingInWholePoints
        //ExFor:CompatibilityOptions.SplitPgBreakAndParaMark
        //ExFor:CompatibilityOptions.SubFontBySize
        //ExFor:CompatibilityOptions.SuppressBottomSpacing
        //ExFor:CompatibilityOptions.SuppressSpacingAtTopOfPage
        //ExFor:CompatibilityOptions.SuppressSpBfAfterPgBrk
        //ExFor:CompatibilityOptions.SuppressTopSpacing
        //ExFor:CompatibilityOptions.SuppressTopSpacingWP
        //ExFor:CompatibilityOptions.SwapBordersFacingPgs
        //ExFor:CompatibilityOptions.SwapInsideAndOutsideForMirrorIndentsAndRelativePositioning
        //ExFor:CompatibilityOptions.TransparentMetafiles
        //ExFor:CompatibilityOptions.TruncateFontHeightsLikeWP6
        //ExFor:CompatibilityOptions.UICompat97To2003
        //ExFor:CompatibilityOptions.UlTrailSpace
        //ExFor:CompatibilityOptions.UnderlineTabInNumList
        //ExFor:CompatibilityOptions.UseAltKinsokuLineBreakRules
        //ExFor:CompatibilityOptions.UseAnsiKerningPairs
        //ExFor:CompatibilityOptions.UseFELayout
        //ExFor:CompatibilityOptions.UseNormalStyleForList
        //ExFor:CompatibilityOptions.UsePrinterMetrics
        //ExFor:CompatibilityOptions.UseSingleBorderforContiguousCells
        //ExFor:CompatibilityOptions.UseWord2002TableStyleRules
        //ExFor:CompatibilityOptions.UseWord2010TableStyleRules
        //ExFor:CompatibilityOptions.UseWord97LineBreakRules
        //ExFor:CompatibilityOptions.WPJustification
        //ExFor:CompatibilityOptions.WPSpaceWidth
        //ExFor:CompatibilityOptions.WrapTrailSpaces
        //ExSummary:Shows how to optimize the document for different versions of Microsoft Word.
        [Test] //ExSkip
        public void OptimizeFor()
        {
            Document doc = new Document();

            // This object contains an extensive list of flags unique to each document
            // that allow us to facilitate backward compatibility with older versions of Microsoft Word.
            CompatibilityOptions options = doc.CompatibilityOptions;

            // Print the default settings for a blank document.
            Console.WriteLine("\nDefault optimization settings:");
            PrintCompatibilityOptions(options);

            // We can access these settings in Microsoft Word via "File" -> "Options" -> "Advanced" -> "Compatibility options for...".
            doc.Save(ArtifactsDir + "CompatibilityOptions.OptimizeFor.DefaultSettings.docx");

            // We can use the OptimizeFor method to ensure optimal compatibility with a specific Microsoft Word version.
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);
            Console.WriteLine("\nOptimized for Word 2010:");
            PrintCompatibilityOptions(options);

            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2000);
            Console.WriteLine("\nOptimized for Word 2000:");
            PrintCompatibilityOptions(options);
        }

        /// <summary>
        /// Groups all flags in a document's compatibility options object by state, then prints each group.
        /// </summary>
        private static void PrintCompatibilityOptions(CompatibilityOptions options)
        {
            IList<string> enabledOptions = new List<string>();
            IList<string> disabledOptions = new List<string>();
            AddOptionName(options.AdjustLineHeightInTable, "AdjustLineHeightInTable", enabledOptions, disabledOptions);
            AddOptionName(options.AlignTablesRowByRow, "AlignTablesRowByRow", enabledOptions, disabledOptions);
            AddOptionName(options.AllowSpaceOfSameStyleInTable, "AllowSpaceOfSameStyleInTable", enabledOptions, disabledOptions);
            AddOptionName(options.ApplyBreakingRules, "ApplyBreakingRules", enabledOptions, disabledOptions);
            AddOptionName(options.AutoSpaceLikeWord95, "AutoSpaceLikeWord95", enabledOptions, disabledOptions);
            AddOptionName(options.AutofitToFirstFixedWidthCell, "AutofitToFirstFixedWidthCell", enabledOptions, disabledOptions);
            AddOptionName(options.BalanceSingleByteDoubleByteWidth, "BalanceSingleByteDoubleByteWidth", enabledOptions, disabledOptions);
            AddOptionName(options.CachedColBalance, "CachedColBalance", enabledOptions, disabledOptions);
            AddOptionName(options.ConvMailMergeEsc, "ConvMailMergeEsc", enabledOptions, disabledOptions);
            AddOptionName(options.DisableOpenTypeFontFormattingFeatures, "DisableOpenTypeFontFormattingFeatures", enabledOptions, disabledOptions);
            AddOptionName(options.DisplayHangulFixedWidth, "DisplayHangulFixedWidth", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotAutofitConstrainedTables, "DoNotAutofitConstrainedTables", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotBreakConstrainedForcedTable, "DoNotBreakConstrainedForcedTable", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotBreakWrappedTables, "DoNotBreakWrappedTables", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotExpandShiftReturn, "DoNotExpandShiftReturn", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotLeaveBackslashAlone, "DoNotLeaveBackslashAlone", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotSnapToGridInCell, "DoNotSnapToGridInCell", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotSuppressIndentation, "DoNotSnapToGridInCell", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotSuppressParagraphBorders, "DoNotSuppressParagraphBorders", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotUseEastAsianBreakRules, "DoNotUseEastAsianBreakRules", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotUseHTMLParagraphAutoSpacing, "DoNotUseHTMLParagraphAutoSpacing", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotUseIndentAsNumberingTabStop, "DoNotUseIndentAsNumberingTabStop", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotVertAlignCellWithSp, "DoNotVertAlignCellWithSp", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotVertAlignInTxbx, "DoNotVertAlignInTxbx", enabledOptions, disabledOptions);
            AddOptionName(options.DoNotWrapTextWithPunct, "DoNotWrapTextWithPunct", enabledOptions, disabledOptions);
            AddOptionName(options.FootnoteLayoutLikeWW8, "FootnoteLayoutLikeWW8", enabledOptions, disabledOptions);
            AddOptionName(options.ForgetLastTabAlignment, "ForgetLastTabAlignment", enabledOptions, disabledOptions);
            AddOptionName(options.GrowAutofit, "GrowAutofit", enabledOptions, disabledOptions);
            AddOptionName(options.LayoutRawTableWidth, "LayoutRawTableWidth", enabledOptions, disabledOptions);
            AddOptionName(options.LayoutTableRowsApart, "LayoutTableRowsApart", enabledOptions, disabledOptions);
            AddOptionName(options.LineWrapLikeWord6, "LineWrapLikeWord6", enabledOptions, disabledOptions);
            AddOptionName(options.MWSmallCaps, "MWSmallCaps", enabledOptions, disabledOptions);
            AddOptionName(options.NoColumnBalance, "NoColumnBalance", enabledOptions, disabledOptions);
            AddOptionName(options.NoExtraLineSpacing, "NoExtraLineSpacing", enabledOptions, disabledOptions);
            AddOptionName(options.NoLeading, "NoLeading", enabledOptions, disabledOptions);
            AddOptionName(options.NoSpaceRaiseLower, "NoSpaceRaiseLower", enabledOptions, disabledOptions);
            AddOptionName(options.NoTabHangInd, "NoTabHangInd", enabledOptions, disabledOptions);
            AddOptionName(options.OverrideTableStyleFontSizeAndJustification, "OverrideTableStyleFontSizeAndJustification", enabledOptions, disabledOptions);
            AddOptionName(options.PrintBodyTextBeforeHeader, "PrintBodyTextBeforeHeader", enabledOptions, disabledOptions);
            AddOptionName(options.PrintColBlack, "PrintColBlack", enabledOptions, disabledOptions);
            AddOptionName(options.SelectFldWithFirstOrLastChar, "SelectFldWithFirstOrLastChar", enabledOptions, disabledOptions);
            AddOptionName(options.ShapeLayoutLikeWW8, "ShapeLayoutLikeWW8", enabledOptions, disabledOptions);
            AddOptionName(options.ShowBreaksInFrames, "ShowBreaksInFrames", enabledOptions, disabledOptions);
            AddOptionName(options.SpaceForUL, "SpaceForUL", enabledOptions, disabledOptions);
            AddOptionName(options.SpacingInWholePoints, "SpacingInWholePoints", enabledOptions, disabledOptions);
            AddOptionName(options.SplitPgBreakAndParaMark, "SplitPgBreakAndParaMark", enabledOptions, disabledOptions);
            AddOptionName(options.SubFontBySize, "SubFontBySize", enabledOptions, disabledOptions);
            AddOptionName(options.SuppressBottomSpacing, "SuppressBottomSpacing", enabledOptions, disabledOptions);
            AddOptionName(options.SuppressSpBfAfterPgBrk, "SuppressSpBfAfterPgBrk", enabledOptions, disabledOptions);
            AddOptionName(options.SuppressSpacingAtTopOfPage, "SuppressSpacingAtTopOfPage", enabledOptions, disabledOptions);
            AddOptionName(options.SuppressTopSpacing, "SuppressTopSpacing", enabledOptions, disabledOptions);
            AddOptionName(options.SuppressTopSpacingWP, "SuppressTopSpacingWP", enabledOptions, disabledOptions);
            AddOptionName(options.SwapBordersFacingPgs, "SwapBordersFacingPgs", enabledOptions, disabledOptions);
            AddOptionName(options.SwapInsideAndOutsideForMirrorIndentsAndRelativePositioning, "SwapInsideAndOutsideForMirrorIndentsAndRelativePositioning", enabledOptions, disabledOptions);
            AddOptionName(options.TransparentMetafiles, "TransparentMetafiles", enabledOptions, disabledOptions);
            AddOptionName(options.TruncateFontHeightsLikeWP6, "TruncateFontHeightsLikeWP6", enabledOptions, disabledOptions);
            AddOptionName(options.UICompat97To2003, "UICompat97To2003", enabledOptions, disabledOptions);
            AddOptionName(options.UlTrailSpace, "UlTrailSpace", enabledOptions, disabledOptions);
            AddOptionName(options.UnderlineTabInNumList, "UnderlineTabInNumList", enabledOptions, disabledOptions);
            AddOptionName(options.UseAltKinsokuLineBreakRules, "UseAltKinsokuLineBreakRules", enabledOptions, disabledOptions);
            AddOptionName(options.UseAnsiKerningPairs, "UseAnsiKerningPairs", enabledOptions, disabledOptions);
            AddOptionName(options.UseFELayout, "UseFELayout", enabledOptions, disabledOptions);
            AddOptionName(options.UseNormalStyleForList, "UseNormalStyleForList", enabledOptions, disabledOptions);
            AddOptionName(options.UsePrinterMetrics, "UsePrinterMetrics", enabledOptions, disabledOptions);
            AddOptionName(options.UseSingleBorderforContiguousCells, "UseSingleBorderforContiguousCells", enabledOptions, disabledOptions);
            AddOptionName(options.UseWord2002TableStyleRules, "UseWord2002TableStyleRules", enabledOptions, disabledOptions);
            AddOptionName(options.UseWord2010TableStyleRules, "UseWord2010TableStyleRules", enabledOptions, disabledOptions);
            AddOptionName(options.UseWord97LineBreakRules, "UseWord97LineBreakRules", enabledOptions, disabledOptions);
            AddOptionName(options.WPJustification, "WPJustification", enabledOptions, disabledOptions);
            AddOptionName(options.WPSpaceWidth, "WPSpaceWidth", enabledOptions, disabledOptions);
            AddOptionName(options.WrapTrailSpaces, "WrapTrailSpaces", enabledOptions, disabledOptions);
            Console.WriteLine("\tEnabled options:");
            foreach (string optionName in enabledOptions)
                Console.WriteLine($"\t\t{optionName}");
            Console.WriteLine("\tDisabled options:");
            foreach (string optionName in disabledOptions)
                Console.WriteLine($"\t\t{optionName}");
        }

        private static void AddOptionName(Boolean option, String optionName, IList<string> enabledOptions, IList<string> disabledOptions)
        {
            if (option)
                enabledOptions.Add(optionName);
            else
                disabledOptions.Add(optionName);
        }
        //ExEnd

        [Test]
        public void Tables()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2002);

            Assert.That(compatibilityOptions.AdjustLineHeightInTable, Is.EqualTo(false));
            Assert.That(compatibilityOptions.AlignTablesRowByRow, Is.EqualTo(false));
            Assert.That(compatibilityOptions.AllowSpaceOfSameStyleInTable, Is.EqualTo(true));
            Assert.That(compatibilityOptions.DoNotAutofitConstrainedTables, Is.EqualTo(true));
            Assert.That(compatibilityOptions.DoNotBreakConstrainedForcedTable, Is.EqualTo(true));
            Assert.That(compatibilityOptions.DoNotBreakWrappedTables, Is.EqualTo(false));
            Assert.That(compatibilityOptions.DoNotSnapToGridInCell, Is.EqualTo(false));
            Assert.That(compatibilityOptions.DoNotUseHTMLParagraphAutoSpacing, Is.EqualTo(false));
            Assert.That(compatibilityOptions.DoNotVertAlignCellWithSp, Is.EqualTo(true));
            Assert.That(compatibilityOptions.ForgetLastTabAlignment, Is.EqualTo(false));
            Assert.That(compatibilityOptions.GrowAutofit, Is.EqualTo(true));
            Assert.That(compatibilityOptions.LayoutRawTableWidth, Is.EqualTo(false));
            Assert.That(compatibilityOptions.LayoutTableRowsApart, Is.EqualTo(false));
            Assert.That(compatibilityOptions.NoColumnBalance, Is.EqualTo(false));
            Assert.That(compatibilityOptions.OverrideTableStyleFontSizeAndJustification, Is.EqualTo(false));
            Assert.That(compatibilityOptions.UseSingleBorderforContiguousCells, Is.EqualTo(false));
            Assert.That(compatibilityOptions.UseWord2002TableStyleRules, Is.EqualTo(true));
            Assert.That(compatibilityOptions.UseWord2010TableStyleRules, Is.EqualTo(false));

            // In the output document, these settings can be accessed in Microsoft Word via
            // File -> Options -> Advanced -> Compatibility options for...
            doc.Save(ArtifactsDir + "CompatibilityOptions.Tables.docx");
        }

        [Test]
        public void Breaks()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.That(compatibilityOptions.ApplyBreakingRules, Is.EqualTo(false));
            Assert.That(compatibilityOptions.DoNotUseEastAsianBreakRules, Is.EqualTo(true));
            Assert.That(compatibilityOptions.ShowBreaksInFrames, Is.EqualTo(false));
            Assert.That(compatibilityOptions.SplitPgBreakAndParaMark, Is.EqualTo(true));
            Assert.That(compatibilityOptions.UseAltKinsokuLineBreakRules, Is.EqualTo(true));
            Assert.That(compatibilityOptions.UseWord97LineBreakRules, Is.EqualTo(false));

            // In the output document, these settings can be accessed in Microsoft Word via
            // File -> Options -> Advanced -> Compatibility options for...
            doc.Save(ArtifactsDir + "CompatibilityOptions.Breaks.docx");
        }

        [Test]
        public void Spacing()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.That(compatibilityOptions.AutoSpaceLikeWord95, Is.EqualTo(false));
            Assert.That(compatibilityOptions.DisplayHangulFixedWidth, Is.EqualTo(true));
            Assert.That(compatibilityOptions.NoExtraLineSpacing, Is.EqualTo(false));
            Assert.That(compatibilityOptions.NoLeading, Is.EqualTo(false));
            Assert.That(compatibilityOptions.NoSpaceRaiseLower, Is.EqualTo(false));
            Assert.That(compatibilityOptions.SpaceForUL, Is.EqualTo(false));
            Assert.That(compatibilityOptions.SpacingInWholePoints, Is.EqualTo(false));
            Assert.That(compatibilityOptions.SuppressBottomSpacing, Is.EqualTo(false));
            Assert.That(compatibilityOptions.SuppressSpBfAfterPgBrk, Is.EqualTo(false));
            Assert.That(compatibilityOptions.SuppressSpacingAtTopOfPage, Is.EqualTo(false));
            Assert.That(compatibilityOptions.SuppressTopSpacing, Is.EqualTo(false));
            Assert.That(compatibilityOptions.UlTrailSpace, Is.EqualTo(false));

            // In the output document, these settings can be accessed in Microsoft Word via
            // File -> Options -> Advanced -> Compatibility options for...
            doc.Save(ArtifactsDir + "CompatibilityOptions.Spacing.docx");
        }

        [Test]
        public void WordPerfect()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.That(compatibilityOptions.SuppressTopSpacingWP, Is.EqualTo(false));
            Assert.That(compatibilityOptions.TruncateFontHeightsLikeWP6, Is.EqualTo(false));
            Assert.That(compatibilityOptions.WPJustification, Is.EqualTo(false));
            Assert.That(compatibilityOptions.WPSpaceWidth, Is.EqualTo(false));
            Assert.That(compatibilityOptions.WrapTrailSpaces, Is.EqualTo(false));

            // In the output document, these settings can be accessed in Microsoft Word via
            // File -> Options -> Advanced -> Compatibility options for...
            doc.Save(ArtifactsDir + "CompatibilityOptions.WordPerfect.docx");
        }

        [Test]
        public void Alignment()
        {
            Document doc = new Document();
            
            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.That(compatibilityOptions.CachedColBalance, Is.EqualTo(true));
            Assert.That(compatibilityOptions.DoNotVertAlignInTxbx, Is.EqualTo(true));
            Assert.That(compatibilityOptions.DoNotWrapTextWithPunct, Is.EqualTo(true));
            Assert.That(compatibilityOptions.NoTabHangInd, Is.EqualTo(false));

            // In the output document, these settings can be accessed in Microsoft Word via
            // File -> Options -> Advanced -> Compatibility options for...
            doc.Save(ArtifactsDir + "CompatibilityOptions.Alignment.docx");
        }

        [Test]
        public void Legacy()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.That(compatibilityOptions.FootnoteLayoutLikeWW8, Is.EqualTo(false));
            Assert.That(compatibilityOptions.LineWrapLikeWord6, Is.EqualTo(false));
            Assert.That(compatibilityOptions.MWSmallCaps, Is.EqualTo(false));
            Assert.That(compatibilityOptions.ShapeLayoutLikeWW8, Is.EqualTo(false));
            Assert.That(compatibilityOptions.UICompat97To2003, Is.EqualTo(false));

            // In the output document, these settings can be accessed in Microsoft Word via
            // File -> Options -> Advanced -> Compatibility options for...
            doc.Save(ArtifactsDir + "CompatibilityOptions.Legacy.docx");
        }

        [Test]
        public void List()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.That(compatibilityOptions.UnderlineTabInNumList, Is.EqualTo(true));
            Assert.That(compatibilityOptions.UseNormalStyleForList, Is.EqualTo(true));

            // In the output document, these settings can be accessed in Microsoft Word via
            // File -> Options -> Advanced -> Compatibility options for...
            doc.Save(ArtifactsDir + "CompatibilityOptions.List.docx");
        }

        [Test]
        public void Misc()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.That(compatibilityOptions.BalanceSingleByteDoubleByteWidth, Is.EqualTo(false));
            Assert.That(compatibilityOptions.ConvMailMergeEsc, Is.EqualTo(false));
            Assert.That(compatibilityOptions.DoNotExpandShiftReturn, Is.EqualTo(false));
            Assert.That(compatibilityOptions.DoNotLeaveBackslashAlone, Is.EqualTo(false));
            Assert.That(compatibilityOptions.DoNotSuppressParagraphBorders, Is.EqualTo(false));
            Assert.That(compatibilityOptions.DoNotUseIndentAsNumberingTabStop, Is.EqualTo(true));
            Assert.That(compatibilityOptions.PrintBodyTextBeforeHeader, Is.EqualTo(false));
            Assert.That(compatibilityOptions.PrintColBlack, Is.EqualTo(false));
            Assert.That(compatibilityOptions.SelectFldWithFirstOrLastChar, Is.EqualTo(true));
            Assert.That(compatibilityOptions.SubFontBySize, Is.EqualTo(false));
            Assert.That(compatibilityOptions.SwapBordersFacingPgs, Is.EqualTo(false));
            Assert.That(compatibilityOptions.TransparentMetafiles, Is.EqualTo(false));
            Assert.That(compatibilityOptions.UseAnsiKerningPairs, Is.EqualTo(true));
            Assert.That(compatibilityOptions.UseFELayout, Is.EqualTo(false));
            Assert.That(compatibilityOptions.UsePrinterMetrics, Is.EqualTo(false));

            // In the output document, these settings can be accessed in Microsoft Word via
            // File -> Options -> Advanced -> Compatibility options for...
            doc.Save(ArtifactsDir + "CompatibilityOptions.Misc.docx");
        }
    }
}