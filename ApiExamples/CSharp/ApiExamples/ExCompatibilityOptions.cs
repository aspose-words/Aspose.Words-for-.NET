// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
        //ExSummary:Shows how to optimize document for different word versions.
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
        private static void PrintCompatibilityOptions(CompatibilityOptions options)
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
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2002);

            Assert.AreEqual(false, compatibilityOptions.AdjustLineHeightInTable);
            Assert.AreEqual(false, compatibilityOptions.AlignTablesRowByRow);
            Assert.AreEqual(true, compatibilityOptions.AllowSpaceOfSameStyleInTable);
            Assert.AreEqual(true, compatibilityOptions.DoNotAutofitConstrainedTables);
            Assert.AreEqual(true, compatibilityOptions.DoNotBreakConstrainedForcedTable);
            Assert.AreEqual(false, compatibilityOptions.DoNotBreakWrappedTables);
            Assert.AreEqual(false, compatibilityOptions.DoNotSnapToGridInCell);
            Assert.AreEqual(false, compatibilityOptions.DoNotUseHTMLParagraphAutoSpacing);
            Assert.AreEqual(true, compatibilityOptions.DoNotVertAlignCellWithSp);
            Assert.AreEqual(false, compatibilityOptions.ForgetLastTabAlignment);
            Assert.AreEqual(true, compatibilityOptions.GrowAutofit);
            Assert.AreEqual(false, compatibilityOptions.LayoutRawTableWidth);
            Assert.AreEqual(false, compatibilityOptions.LayoutTableRowsApart);
            Assert.AreEqual(false, compatibilityOptions.NoColumnBalance);
            Assert.AreEqual(false, compatibilityOptions.OverrideTableStyleFontSizeAndJustification);
            Assert.AreEqual(false, compatibilityOptions.UseSingleBorderforContiguousCells);
            Assert.AreEqual(true, compatibilityOptions.UseWord2002TableStyleRules);
            Assert.AreEqual(false, compatibilityOptions.UseWord2010TableStyleRules);

            // These options will become available in File > Options > Advanced > Compatibility Options in the output document
            doc.Save(ArtifactsDir + "CompatibilityOptionsTable.docx");
        }

        [Test]
        public void CompatibilityOptionsBreaks()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(false, compatibilityOptions.ApplyBreakingRules);
            Assert.AreEqual(true, compatibilityOptions.DoNotUseEastAsianBreakRules);
            Assert.AreEqual(false, compatibilityOptions.ShowBreaksInFrames);
            Assert.AreEqual(true, compatibilityOptions.SplitPgBreakAndParaMark);
            Assert.AreEqual(true, compatibilityOptions.UseAltKinsokuLineBreakRules);
            Assert.AreEqual(false, compatibilityOptions.UseWord97LineBreakRules);

            // These options will become available in File > Options > Advanced > Compatibility Options in the output document
            doc.Save(ArtifactsDir + "CompatibilityOptionsBreaks.docx");
        }

        [Test]
        public void CompatibilityOptionsSpacing()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(false, compatibilityOptions.AutoSpaceLikeWord95);
            Assert.AreEqual(true, compatibilityOptions.DisplayHangulFixedWidth);
            Assert.AreEqual(false, compatibilityOptions.NoExtraLineSpacing);
            Assert.AreEqual(false, compatibilityOptions.NoLeading);
            Assert.AreEqual(false, compatibilityOptions.NoSpaceRaiseLower);
            Assert.AreEqual(false, compatibilityOptions.SpaceForUL);
            Assert.AreEqual(false, compatibilityOptions.SpacingInWholePoints);
            Assert.AreEqual(false, compatibilityOptions.SuppressBottomSpacing);
            Assert.AreEqual(false, compatibilityOptions.SuppressSpBfAfterPgBrk);
            Assert.AreEqual(false, compatibilityOptions.SuppressSpacingAtTopOfPage);
            Assert.AreEqual(false, compatibilityOptions.SuppressTopSpacing);
            Assert.AreEqual(false, compatibilityOptions.UlTrailSpace);

            // These options will become available in File > Options > Advanced > Compatibility Options in the output document
            doc.Save(ArtifactsDir + "CompatibilityOptionsSpacing.docx");
        }

        [Test]
        public void CompatibilityOptionsWordPerfect()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(false, compatibilityOptions.SuppressTopSpacingWP);
            Assert.AreEqual(false, compatibilityOptions.TruncateFontHeightsLikeWP6);
            Assert.AreEqual(false, compatibilityOptions.WPJustification);
            Assert.AreEqual(false, compatibilityOptions.WPSpaceWidth);
            Assert.AreEqual(false, compatibilityOptions.WrapTrailSpaces);

            // These options will become available in File > Options > Advanced > Compatibility Options in the output document
            doc.Save(ArtifactsDir + "CompatibilityOptionsWordPerfect.docx");
        }

        [Test]
        public void CompatibilityOptionsAlignment()
        {
            Document doc = new Document();
            
            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(true, compatibilityOptions.CachedColBalance);
            Assert.AreEqual(true, compatibilityOptions.DoNotVertAlignInTxbx);
            Assert.AreEqual(true, compatibilityOptions.DoNotWrapTextWithPunct);
            Assert.AreEqual(false, compatibilityOptions.NoTabHangInd);

            // These options will become available in File > Options > Advanced > Compatibility Options in the output document
            doc.Save(ArtifactsDir + "CompatibilityOptionsAlignment.docx");
        }

        [Test]
        public void CompatibilityOptionsLegacy()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(false, compatibilityOptions.FootnoteLayoutLikeWW8);
            Assert.AreEqual(false, compatibilityOptions.LineWrapLikeWord6);
            Assert.AreEqual(false, compatibilityOptions.MWSmallCaps);
            Assert.AreEqual(false, compatibilityOptions.ShapeLayoutLikeWW8);
            Assert.AreEqual(false, compatibilityOptions.UICompat97To2003);

            // These options will become available in File > Options > Advanced > Compatibility Options in the output document
            doc.Save(ArtifactsDir + "CompatibilityOptionsLegacy.docx");
        }

        [Test]
        public void CompatibilityOptionsList()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(true, compatibilityOptions.UnderlineTabInNumList);
            Assert.AreEqual(true, compatibilityOptions.UseNormalStyleForList);

            // These options will become available in File > Options > Advanced > Compatibility Options in the output document
            doc.Save(ArtifactsDir + "CompatibilityOptionsList.docx");
        }

        [Test]
        public void CompatibilityOptionsMisc()
        {
            Document doc = new Document();

            CompatibilityOptions compatibilityOptions = doc.CompatibilityOptions;
            compatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(false, compatibilityOptions.BalanceSingleByteDoubleByteWidth);
            Assert.AreEqual(false, compatibilityOptions.ConvMailMergeEsc);
            Assert.AreEqual(false, compatibilityOptions.DoNotExpandShiftReturn);
            Assert.AreEqual(false, compatibilityOptions.DoNotLeaveBackslashAlone);
            Assert.AreEqual(false, compatibilityOptions.DoNotSuppressParagraphBorders);
            Assert.AreEqual(true, compatibilityOptions.DoNotUseIndentAsNumberingTabStop);
            Assert.AreEqual(false, compatibilityOptions.PrintBodyTextBeforeHeader);
            Assert.AreEqual(false, compatibilityOptions.PrintColBlack);
            Assert.AreEqual(true, compatibilityOptions.SelectFldWithFirstOrLastChar);
            Assert.AreEqual(false, compatibilityOptions.SubFontBySize);
            Assert.AreEqual(false, compatibilityOptions.SwapBordersFacingPgs);
            Assert.AreEqual(false, compatibilityOptions.TransparentMetafiles);
            Assert.AreEqual(true, compatibilityOptions.UseAnsiKerningPairs);
            Assert.AreEqual(false, compatibilityOptions.UseFELayout);
            Assert.AreEqual(false, compatibilityOptions.UsePrinterMetrics);

            // These options will become available in File > Options > Advanced > Compatibility Options in the output document
            doc.Save(ArtifactsDir + "CompatibilityOptionsMisc.docx");
        }
    }
}