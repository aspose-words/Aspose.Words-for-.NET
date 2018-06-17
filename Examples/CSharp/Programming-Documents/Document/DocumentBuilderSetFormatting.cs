﻿using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderSetFormatting
    {
        public static void Run()
        {
            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            SetFontFormatting(dataDir);
            SetParagraphFormatting(dataDir);
            SetTableCellFormatting(dataDir);
            SetMultilevelListFormatting(dataDir);
            SetPageSetupAndSectionFormatting(dataDir);
            ApplyParagraphStyle(dataDir);
            ApplyBordersAndShadingToParagraph(dataDir);
            SetAsianTypographyLinebreakGroupProp(dataDir);
        }

        public static void SetSpacebetweenAsianandLatintext(string dataDir)
        {
            // ExStart:DocumentBuilderSetSpacebetweenAsianandLatintext
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph formatting properties
            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.AddSpaceBetweenFarEastAndAlpha = true;
            paragraphFormat.AddSpaceBetweenFarEastAndDigit = true;

            builder.Writeln("Automatically adjust space between Asian and Latin text");
            builder.Writeln("Automatically adjust space between Asian text and numbers");

            dataDir = dataDir + "DocumentBuilderSetSpacebetweenAsianandLatintext.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentBuilderSetSpacebetweenAsianandLatintext
            Console.WriteLine("\nParagraphFormat properties AddSpaceBetweenFarEastAndAlpha and AddSpaceBetweenFarEastAndDigit set successfully.\nFile saved at " + dataDir);
        }

        public static void SetAsianTypographyLinebreakGroupProp(string dataDir)
        {
            // ExStart:SetAsianTypographyLinebreakGroupProp
            Document doc = new Document(dataDir + "Input.docx");

            ParagraphFormat format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;
            format.FarEastLineBreakControl = false;
            format.WordWrap = true;
            format.HangingPunctuation = false;

            dataDir = dataDir + "SetAsianTypographyLinebreakGroupProp_out.docx";
            doc.Save(dataDir);
            // ExEnd:SetAsianTypographyLinebreakGroupProp
            Console.WriteLine("\nParagraphFormat properties for Asian Typography line break group are set successfully.\nFile saved at " + dataDir);
        }

        public static void SetFontFormatting(string dataDir)
        {
            // ExStart:DocumentBuilderSetFontFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set font formatting properties
            Font font = builder.Font;
            font.Bold = true;
            font.Color = System.Drawing.Color.DarkBlue;
            font.Italic = true;
            font.Name = "Arial";
            font.Size = 24;
            font.Spacing = 5;
            font.Underline = Underline.Double;

            // Output formatted text
            builder.Writeln("I'm a very nice formatted string.");
            dataDir = dataDir + "DocumentBuilderSetFontFormatting_out.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentBuilderSetFontFormatting
            Console.WriteLine("\nFont formatting using DocumentBuilder set successfully.\nFile saved at " + dataDir);
        }
        public static void SetParagraphFormatting(string dataDir)
        {
            // ExStart:DocumentBuilderSetParagraphFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph formatting properties
            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.Alignment = ParagraphAlignment.Center;
            paragraphFormat.LeftIndent = 50;
            paragraphFormat.RightIndent = 50;
            paragraphFormat.SpaceAfter = 25;

            // Output text
            builder.Writeln("I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
            builder.Writeln("I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");

            dataDir = dataDir + "DocumentBuilderSetParagraphFormatting_out.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentBuilderSetParagraphFormatting
            Console.WriteLine("\nParagraph formatting using DocumentBuilder set successfully.\nFile saved at " + dataDir);
        }
        public static void SetTableCellFormatting(string dataDir)
        {
            // ExStart:DocumentBuilderSetTableCellFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.InsertCell();

            // Set the cell formatting
            CellFormat cellFormat = builder.CellFormat;
            cellFormat.Width = 250;
            cellFormat.LeftPadding = 30;
            cellFormat.RightPadding = 30;
            cellFormat.TopPadding = 30;
            cellFormat.BottomPadding = 30;

            builder.Writeln("I'm a wonderful formatted cell.");

            builder.EndRow();
            builder.EndTable();

            dataDir = dataDir + "DocumentBuilderSetTableCellFormatting_out.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentBuilderSetTableCellFormatting
            Console.WriteLine("\nTable cell formatting using DocumentBuilder set successfully.\nFile saved at " + dataDir);
        }
        public static void SetTableRowFormatting(string dataDir)
        {
            // ExStart:DocumentBuilderSetTableRowFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();

            // Set the row formatting
            RowFormat rowFormat = builder.RowFormat;
            rowFormat.Height = 100;
            rowFormat.HeightRule = HeightRule.Exactly;
            // These formatting properties are set on the table and are applied to all rows in the table.
            table.LeftPadding = 30;
            table.RightPadding = 30;
            table.TopPadding = 30;
            table.BottomPadding = 30;

            builder.Writeln("I'm a wonderful formatted row.");

            builder.EndRow();
            builder.EndTable();

            dataDir = dataDir + "DocumentBuilderSetTableRowFormatting_out.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentBuilderSetTableRowFormatting
            Console.WriteLine("\nTable row formatting using DocumentBuilder set successfully.\nFile saved at " + dataDir);
        }
        public static void SetMultilevelListFormatting(string dataDir)
        {
            // ExStart:DocumentBuilderSetMultilevelListFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ListFormat.ApplyNumberDefault();

            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            builder.ListFormat.ListIndent();

            builder.Writeln("Item 2.1");
            builder.Writeln("Item 2.2");

            builder.ListFormat.ListIndent();

            builder.Writeln("Item 2.2.1");
            builder.Writeln("Item 2.2.2");

            builder.ListFormat.ListOutdent();

            builder.Writeln("Item 2.3");

            builder.ListFormat.ListOutdent();

            builder.Writeln("Item 3");

            builder.ListFormat.RemoveNumbers();
            dataDir = dataDir + "DocumentBuilderSetMultilevelListFormatting_out.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentBuilderSetMultilevelListFormatting
            Console.WriteLine("\nMultilevel list formatting using DocumentBuilder set successfully.\nFile saved at " + dataDir);
        }
        public static void SetPageSetupAndSectionFormatting(string dataDir)
        {
            // ExStart:DocumentBuilderSetPageSetupAndSectionFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set page properties
            builder.PageSetup.Orientation = Orientation.Landscape;
            builder.PageSetup.LeftMargin = 50;
            builder.PageSetup.PaperSize = PaperSize.Paper10x14;

            dataDir = dataDir + "DocumentBuilderSetPageSetupAndSectionFormatting_out.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentBuilderSetPageSetupAndSectionFormatting
            Console.WriteLine("\nPage setup and section formatting using DocumentBuilder set successfully.\nFile saved at " + dataDir);
        }
        public static void ApplyParagraphStyle(string dataDir)
        {
            // ExStart:DocumentBuilderApplyParagraphStyle
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph style
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;

            builder.Write("Hello");
            dataDir = dataDir + "DocumentBuilderApplyParagraphStyle_out.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentBuilderApplyParagraphStyle
            Console.WriteLine("\nParagraph style using DocumentBuilder applied successfully.\nFile saved at " + dataDir);
        }
        public static void ApplyBordersAndShadingToParagraph(string dataDir)
        {
            // ExStart:DocumentBuilderApplyBordersAndShadingToParagraph
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph borders
            BorderCollection borders = builder.ParagraphFormat.Borders;
            borders.DistanceFromText = 20;
            borders[BorderType.Left].LineStyle = LineStyle.Double;
            borders[BorderType.Right].LineStyle = LineStyle.Double;
            borders[BorderType.Top].LineStyle = LineStyle.Double;
            borders[BorderType.Bottom].LineStyle = LineStyle.Double;

            // Set paragraph shading
            Shading shading = builder.ParagraphFormat.Shading;
            shading.Texture = TextureIndex.TextureDiagonalCross;
            shading.BackgroundPatternColor = System.Drawing.Color.LightCoral;
            shading.ForegroundPatternColor = System.Drawing.Color.LightSalmon;

            builder.Write("I'm a formatted paragraph with double border and nice shading.");
            dataDir = dataDir + "DocumentBuilderApplyBordersAndShadingToParagraph_out.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentBuilderApplyBordersAndShadingToParagraph
            Console.WriteLine("\nBorders and shading using DocumentBuilder applied successfully to paragraph.\nFile saved at " + dataDir);
        }
        
    }
}
