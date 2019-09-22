// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
        //ExSummary:Shows how to optimize our document for different word versions.
        [Test] //ExSkip
        public void CompatibilityOptionsOptimizeFor()
        {
            // Create a blank document and get its CompatibilityOptions object
            Document doc = new Document();
            CompatibilityOptions options = doc.CompatibilityOptions;

            // By default, the CompatibilityOptions will contain the set of values printed below
            Console.WriteLine("\nDefault optimization settings:");
            PrintCompatibilityOptions(options);

            // These attributes can be accessed in the output document via File > Options > Advanced > Compatibility for...
            doc.Save(ArtifactsDir + "DefaultCompatibility.docx");

            // We can use the OptimizeFor method to set these values automatically
            // for maximum compatibility with some Microsoft Word versions
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);
            Console.WriteLine("\nOptimized for Word 2010:");
            PrintCompatibilityOptions(options);

            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2000);
            Console.WriteLine("\nOptimized for Word 2000:");
            PrintCompatibilityOptions(options);
        }

        /// <summary>
        /// Prints all options of a CompatibilityOptions object and indicates whether they are enabled or disabled
        /// </summary>
        private void PrintCompatibilityOptions(CompatibilityOptions options)
        {
            for (int i = 1; i >= 0; i--)
            {
                Console.WriteLine(Convert.ToBoolean(i) ? "\tEnabled options:" : "\tDisabled options:");
                SortedSet<string> optionNames = new SortedSet<string>();

                foreach (System.ComponentModel.PropertyDescriptor descriptor in System.ComponentModel.TypeDescriptor.GetProperties(options))
                {
                    if (descriptor.PropertyType == Type.GetType("System.Boolean") && i == Convert.ToInt32(descriptor.GetValue(options)))
                    {
                        optionNames.Add(descriptor.Name);
                    }
                }

                foreach (string s in optionNames)
                {
                    Console.WriteLine($"\t\t{s}");
                }
            }
        }
        //ExEnd

        [Test]
        public void CompatibilityOptionsTable()
        {
            //ExStart
            //ExFor:CompatibilityOptions.AdjustLineHeightInTable
            //ExFor:CompatibilityOptions.AlignTablesRowByRow
            //ExFor:CompatibilityOptions.AllowSpaceOfSameStyleInTable
            //ExFor:CompatibilityOptions.DoNotAutofitConstrainedTables
            //ExFor:CompatibilityOptions.DoNotBreakConstrainedForcedTable
            //ExFor:CompatibilityOptions.DoNotBreakWrappedTables
            //ExFor:CompatibilityOptions.DoNotSnapToGridInCell
            //ExFor:CompatibilityOptions.DoNotUseHTMLParagraphAutoSpacing
            //ExFor:CompatibilityOptions.DoNotVertAlignCellWithSp		
            //ExFor:CompatibilityOptions.ForgetLastTabAlignment
            //ExFor:CompatibilityOptions.GrowAutofit
            //ExFor:CompatibilityOptions.LayoutRawTableWidth
            //ExFor:CompatibilityOptions.LayoutTableRowsApart
            //ExFor:CompatibilityOptions.NoColumnBalance
            //ExFor:CompatibilityOptions.OverrideTableStyleFontSizeAndJustification
            //ExFor:CompatibilityOptions.UseSingleBorderforContiguousCells
            //ExFor:CompatibilityOptions.UseWord2002TableStyleRules
            //ExFor:CompatibilityOptions.UseWord2010TableStyleRules
            //ExSummary:Shows how to set compatibility options pertaining to tables to Word 2002 settings.
            Document doc = new Document();
            CompatibilityOptions co = doc.CompatibilityOptions;

            co.OptimizeFor(MsWordVersion.Word2002);

            Assert.AreEqual(false, co.AdjustLineHeightInTable);
            Assert.AreEqual(false, co.AlignTablesRowByRow);
            Assert.AreEqual(true, co.AllowSpaceOfSameStyleInTable);
            Assert.AreEqual(true, co.DoNotAutofitConstrainedTables);
            Assert.AreEqual(true, co.DoNotBreakConstrainedForcedTable);
            Assert.AreEqual(false, co.DoNotBreakWrappedTables);
            Assert.AreEqual(false, co.DoNotSnapToGridInCell);
            Assert.AreEqual(false, co.DoNotUseHTMLParagraphAutoSpacing);
            Assert.AreEqual(true, co.DoNotVertAlignCellWithSp);
            Assert.AreEqual(false, co.ForgetLastTabAlignment);
            Assert.AreEqual(true, co.GrowAutofit);
            Assert.AreEqual(false, co.LayoutRawTableWidth);
            Assert.AreEqual(false, co.LayoutTableRowsApart);
            Assert.AreEqual(false, co.NoColumnBalance);
            Assert.AreEqual(false, co.OverrideTableStyleFontSizeAndJustification);
            Assert.AreEqual(false, co.UseSingleBorderforContiguousCells);
            Assert.AreEqual(true, co.UseWord2002TableStyleRules);
            Assert.AreEqual(false, co.UseWord2010TableStyleRules);

            doc.Save(ArtifactsDir + "CompatibilityOptionsTable.docx");
            //ExEnd
        }

        [Test]
        public void CompatibilityOptionsBreaks()
        {
            //ExStart
            //ExFor:CompatibilityOptions.ApplyBreakingRules
            //ExFor:CompatibilityOptions.DoNotUseEastAsianBreakRules
            //ExFor:CompatibilityOptions.ShowBreaksInFrames
            //ExFor:CompatibilityOptions.SplitPgBreakAndParaMark
            //ExFor:CompatibilityOptions.UseAltKinsokuLineBreakRules
            //ExFor:CompatibilityOptions.UseWord97LineBreakRules
            //ExSummary:Shows how to set compatibility options pertaining to breaks to Word 2000 settings.
            Document doc = new Document();
            CompatibilityOptions co = doc.CompatibilityOptions;

            co.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(false, co.ApplyBreakingRules);
            Assert.AreEqual(true, co.DoNotUseEastAsianBreakRules);
            Assert.AreEqual(false, co.ShowBreaksInFrames);
            Assert.AreEqual(true, co.SplitPgBreakAndParaMark);
            Assert.AreEqual(true, co.UseAltKinsokuLineBreakRules);
            Assert.AreEqual(false, co.UseWord97LineBreakRules);

            doc.Save(ArtifactsDir + "CompatibilityOptionsBreaks.docx");
            //ExEnd
        }

        [Test]
        public void CompatibilityOptionsSpacing()
        {
            //ExStart
            //ExFor:CompatibilityOptions.AutoSpaceLikeWord95
            //ExFor:CompatibilityOptions.DisplayHangulFixedWidth
            //ExFor:CompatibilityOptions.NoExtraLineSpacing
            //ExFor:CompatibilityOptions.NoLeading
            //ExFor:CompatibilityOptions.NoSpaceRaiseLower
            //ExFor:CompatibilityOptions.SpaceForUL
            //ExFor:CompatibilityOptions.SpacingInWholePoints
            //ExFor:CompatibilityOptions.SuppressBottomSpacing
            //ExFor:CompatibilityOptions.SuppressSpBfAfterPgBrk
            //ExFor:CompatibilityOptions.SuppressSpacingAtTopOfPage
            //ExFor:CompatibilityOptions.SuppressTopSpacing
            //ExFor:CompatibilityOptions.UlTrailSpace
            //ExSummary:Shows how to set compatibility options pertaining to spacing to Word 2000 settings.
            Document doc = new Document();
            CompatibilityOptions co = doc.CompatibilityOptions;

            co.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(false, co.AutoSpaceLikeWord95);
            Assert.AreEqual(true, co.DisplayHangulFixedWidth);
            Assert.AreEqual(false, co.NoExtraLineSpacing);
            Assert.AreEqual(false, co.NoLeading);
            Assert.AreEqual(false, co.NoSpaceRaiseLower);
            Assert.AreEqual(false, co.SpaceForUL);
            Assert.AreEqual(false, co.SpacingInWholePoints);
            Assert.AreEqual(false, co.SuppressBottomSpacing);
            Assert.AreEqual(false, co.SuppressSpBfAfterPgBrk);
            Assert.AreEqual(false, co.SuppressSpacingAtTopOfPage);
            Assert.AreEqual(false, co.SuppressTopSpacing);
            Assert.AreEqual(false, co.UlTrailSpace);

            doc.Save(ArtifactsDir + "CompatibilityOptionsSpacing.docx");
            //ExEnd
        }

        [Test]
        public void CompatibilityOptionsWordPerfect()
        {
            //ExStart
            //ExFor:CompatibilityOptions.SuppressTopSpacingWP
            //ExFor:CompatibilityOptions.TruncateFontHeightsLikeWP6
            //ExFor:CompatibilityOptions.WPJustification
            //ExFor:CompatibilityOptions.WPSpaceWidth
            //ExFor:CompatibilityOptions.WrapTrailSpaces
            //ExSummary:Shows how to set compatibility options to emulate Corel WordPerfect.
            Document doc = new Document();
            CompatibilityOptions co = doc.CompatibilityOptions;

            co.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(false, co.SuppressTopSpacingWP);
            Assert.AreEqual(false, co.TruncateFontHeightsLikeWP6);
            Assert.AreEqual(false, co.WPJustification);
            Assert.AreEqual(false, co.WPSpaceWidth);
            Assert.AreEqual(false, co.WrapTrailSpaces);

            doc.Save(ArtifactsDir + "CompatibilityOptionsWordPerfect.docx");
            //ExEnd
        }

        [Test]
        public void CompatibilityOptionsAlignment()
        {
            //ExStart
            //ExFor:CompatibilityOptions.CachedColBalance
            //ExFor:CompatibilityOptions.DoNotVertAlignInTxbx
            //ExFor:CompatibilityOptions.DoNotWrapTextWithPunct
            //ExFor:CompatibilityOptions.NoTabHangInd
            //ExSummary:Shows how to set compatibility options pertaining to alignment to Word 2000 settings.
            Document doc = new Document();
            CompatibilityOptions co = doc.CompatibilityOptions;

            co.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(true, co.CachedColBalance);
            Assert.AreEqual(true, co.DoNotVertAlignInTxbx);
            Assert.AreEqual(true, co.DoNotWrapTextWithPunct);
            Assert.AreEqual(false, co.NoTabHangInd);

            doc.Save(ArtifactsDir + "CompatibilityOptionsAlignment.docx");
            //ExEnd
        }

        [Test]
        public void CompatibilityOptionsLegacy()
        {
            //ExStart
            //ExFor:CompatibilityOptions.FootnoteLayoutLikeWW8
            //ExFor:CompatibilityOptions.LineWrapLikeWord6
            //ExFor:CompatibilityOptions.MWSmallCaps
            //ExFor:CompatibilityOptions.ShapeLayoutLikeWW8
            //ExFor:CompatibilityOptions.UICompat97To2003
            //ExSummary:Shows how to set compatibility options to emulate older Microsoft Word versions.
            Document doc = new Document();
            CompatibilityOptions co = doc.CompatibilityOptions;

            co.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(false, co.FootnoteLayoutLikeWW8);
            Assert.AreEqual(false, co.LineWrapLikeWord6);
            Assert.AreEqual(false, co.MWSmallCaps);
            Assert.AreEqual(false, co.ShapeLayoutLikeWW8);
            Assert.AreEqual(false, co.UICompat97To2003);

            doc.Save(ArtifactsDir + "CompatibilityOptionsLegacy.docx");
            //ExEnd
        }

        [Test]
        public void CompatibilityOptionsList()
        {
            //ExStart
            //ExFor:CompatibilityOptions.UnderlineTabInNumList
            //ExFor:CompatibilityOptions.UseNormalStyleForList
            //ExSummary:Shows how to set compatibility options pertaining to lists to Word 2000 settings.
            Document doc = new Document();
            CompatibilityOptions co = doc.CompatibilityOptions;

            co.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(true, co.UnderlineTabInNumList);
            Assert.AreEqual(true, co.UseNormalStyleForList);

            doc.Save(ArtifactsDir + "CompatibilityOptionsList.docx");
            //ExEnd
        }

        [Test]
        public void CompatibilityOptionsMisc()
        {
            //ExStart
            //ExFor:CompatibilityOptions.BalanceSingleByteDoubleByteWidth
            //ExFor:CompatibilityOptions.ConvMailMergeEsc
            //ExFor:CompatibilityOptions.DoNotExpandShiftReturn
            //ExFor:CompatibilityOptions.DoNotLeaveBackslashAlone
            //ExFor:CompatibilityOptions.DoNotSuppressParagraphBorders
            //ExFor:CompatibilityOptions.DoNotUseIndentAsNumberingTabStop
            //ExFor:CompatibilityOptions.PrintBodyTextBeforeHeader
            //ExFor:CompatibilityOptions.PrintColBlack
            //ExFor:CompatibilityOptions.SelectFldWithFirstOrLastChar
            //ExFor:CompatibilityOptions.SubFontBySize
            //ExFor:CompatibilityOptions.SwapBordersFacingPgs
            //ExFor:CompatibilityOptions.TransparentMetafiles
            //ExFor:CompatibilityOptions.UseAnsiKerningPairs
            //ExFor:CompatibilityOptions.UseFELayout
            //ExFor:CompatibilityOptions.UsePrinterMetrics
            //ExSummary:Shows how to set miscellaneous compatibility options to Word 2000 settings.
            Document doc = new Document();
            CompatibilityOptions co = doc.CompatibilityOptions;

            co.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(false, co.BalanceSingleByteDoubleByteWidth);
            Assert.AreEqual(false, co.ConvMailMergeEsc);
            Assert.AreEqual(false, co.DoNotExpandShiftReturn);
            Assert.AreEqual(false, co.DoNotLeaveBackslashAlone);
            Assert.AreEqual(false, co.DoNotSuppressParagraphBorders);
            Assert.AreEqual(true, co.DoNotUseIndentAsNumberingTabStop);
            Assert.AreEqual(false, co.PrintBodyTextBeforeHeader);
            Assert.AreEqual(false, co.PrintColBlack);
            Assert.AreEqual(true, co.SelectFldWithFirstOrLastChar);
            Assert.AreEqual(false, co.SubFontBySize);
            Assert.AreEqual(false, co.SwapBordersFacingPgs);
            Assert.AreEqual(false, co.TransparentMetafiles);
            Assert.AreEqual(true, co.UseAnsiKerningPairs);
            Assert.AreEqual(false, co.UseFELayout);
            Assert.AreEqual(false, co.UsePrinterMetrics);

            doc.Save(ArtifactsDir + "CompatibilityOptionsMisc.docx");
            //ExEnd
        }
    }
}