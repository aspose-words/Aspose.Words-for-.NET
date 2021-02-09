// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
        public void Tables()
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

            Assert.AreEqual(false, compatibilityOptions.ApplyBreakingRules);
            Assert.AreEqual(true, compatibilityOptions.DoNotUseEastAsianBreakRules);
            Assert.AreEqual(false, compatibilityOptions.ShowBreaksInFrames);
            Assert.AreEqual(true, compatibilityOptions.SplitPgBreakAndParaMark);
            Assert.AreEqual(true, compatibilityOptions.UseAltKinsokuLineBreakRules);
            Assert.AreEqual(false, compatibilityOptions.UseWord97LineBreakRules);

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

            Assert.AreEqual(false, compatibilityOptions.SuppressTopSpacingWP);
            Assert.AreEqual(false, compatibilityOptions.TruncateFontHeightsLikeWP6);
            Assert.AreEqual(false, compatibilityOptions.WPJustification);
            Assert.AreEqual(false, compatibilityOptions.WPSpaceWidth);
            Assert.AreEqual(false, compatibilityOptions.WrapTrailSpaces);

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

            Assert.AreEqual(true, compatibilityOptions.CachedColBalance);
            Assert.AreEqual(true, compatibilityOptions.DoNotVertAlignInTxbx);
            Assert.AreEqual(true, compatibilityOptions.DoNotWrapTextWithPunct);
            Assert.AreEqual(false, compatibilityOptions.NoTabHangInd);

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

            Assert.AreEqual(false, compatibilityOptions.FootnoteLayoutLikeWW8);
            Assert.AreEqual(false, compatibilityOptions.LineWrapLikeWord6);
            Assert.AreEqual(false, compatibilityOptions.MWSmallCaps);
            Assert.AreEqual(false, compatibilityOptions.ShapeLayoutLikeWW8);
            Assert.AreEqual(false, compatibilityOptions.UICompat97To2003);

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

            Assert.AreEqual(true, compatibilityOptions.UnderlineTabInNumList);
            Assert.AreEqual(true, compatibilityOptions.UseNormalStyleForList);

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

            // In the output document, these settings can be accessed in Microsoft Word via
            // File -> Options -> Advanced -> Compatibility options for...
            doc.Save(ArtifactsDir + "CompatibilityOptions.Misc.docx");
        }
    }
}