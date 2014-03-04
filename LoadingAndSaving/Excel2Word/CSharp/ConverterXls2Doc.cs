// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.Collections;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;


namespace Excel2Word
{
    /// <summary>
    /// This class contains methods that allow to convert Excel workbook to Word document.
    /// 
    /// A main idea is that Excel table can have unlimited witdh, 
    /// that's why we should split this table, how Excel do in print preview.
    /// 
    /// This class demonstrates how you can use Aspose libraries to convert different types of document.
    /// </summary>
    class ConverterXls2Doc
    {
        /// <summary>
        /// Convert Excel workbook to Word document
        /// </summary>
        /// <param name="workbook">Input workbook</param>
        /// <returns>Word document</returns>
        internal Document Convert(Workbook workbook)
        {
            //Create new document
            Document doc = new Document();
            //Create an instance of the  DocumentBuilder class
            DocumentBuilder builder = new DocumentBuilder(doc);
           
            //Every worksheet in Excel workbook is represented as section in Word document
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                //Import PageSetup from Excel file to Word document
                //Orientation can be Portrait or Landscape
                builder.PageSetup.Orientation = ConvertPageOrientation(worksheet.PageSetup.Orientation);
                //Paper size can be A4, A3, Letter, etc.
                builder.PageSetup.PaperSize = ConvertPaperSize(worksheet.PageSetup.PaperSize);
                //Import margins
                builder.PageSetup.LeftMargin = ConvertUtil.InchToPoint(worksheet.PageSetup.LeftMarginInch); // 1cm = 28.35pt
                builder.PageSetup.RightMargin = ConvertUtil.InchToPoint(worksheet.PageSetup.RightMarginInch);
                builder.PageSetup.TopMargin = ConvertUtil.InchToPoint(worksheet.PageSetup.TopMarginInch);
                builder.PageSetup.BottomMargin = ConvertUtil.InchToPoint(worksheet.PageSetup.BottomMarginInch);

                //Get array of Word tables, every table in this array represents a part of Excel worksheet.
                ArrayList partsArray = GetTablePartsArray(worksheet, doc);
                //Insert all tables into the Word document
                foreach (Table table in partsArray)
                {
                    //Insert table
                    builder.CurrentSection.Body.AppendChild(table);
                    //Move coursore to document end
                    builder.MoveToDocumentEnd();
                    //Insert break if table is not last in the collection
                    if (!table.Equals(partsArray[partsArray.Count - 1]))
                    {
                        builder.InsertBreak(BreakType.SectionBreakNewPage);
                    }
                }
                //Insert break if current workseet is not last in the Excwl workbook
                if (!worksheet.Equals(workbook.Worksheets[workbook.Worksheets.Count - 1]) && partsArray.Count != 0)
                {
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
                }
            }

            return doc;
        }

        /// <summary>
        /// This method returns array of Word tables, every table in this array represents a part of Excel worksheet.
        /// </summary>
        /// <param name="excelWorksheet">Input worksheet</param>
        /// <param name="doc">Parent document</param>
        /// <returns>Array of Word tables</returns>
        private ArrayList GetTablePartsArray(Worksheet excelWorksheet, Document doc)
        {
            //Get column index of cell that contains data
            int colCount = excelWorksheet.Cells.MaxColumn + 1;
            //Get row index of cell that contains data
            int rowCount = excelWorksheet.Cells.MaxRow + 1;
            int startColumn = excelWorksheet.Cells.MinColumn;
            int startRow = excelWorksheet.Cells.MinRow;

            //Get area in the worksheet that will be printed
            //Returns something like this "A1:D51" of null
            string excelPrintArea = excelWorksheet.PageSetup.PrintArea;
            if (!string.IsNullOrEmpty(excelPrintArea))
            {
                //Get first cell in the printed area
                string rangeStart = excelPrintArea.Substring(0, excelPrintArea.IndexOf(":"));
                //Get last cell in the printed area
                string rangeEnd = excelPrintArea.Substring(excelPrintArea.IndexOf(":") + 1);
                //Get printed range from worksheet
                Aspose.Cells.Range range = excelWorksheet.Cells.CreateRange(rangeStart, rangeEnd);

                colCount = range.ColumnCount + range.FirstColumn;
                rowCount = range.RowCount + range.FirstRow;
                startColumn = range.FirstColumn;
                startRow = range.FirstRow;
            }

            //Extract objects like Pictures, Charts, etc and store in the HashTable
            //if worksheet contains object that is placed outside the region then resize region (count of rows and columns)
            Aspose.Cells.Range drawRange = ExtractDrawingObjects(excelWorksheet);
            if (drawRange.RowCount > rowCount)
                rowCount = drawRange.RowCount;
            if (drawRange.ColumnCount > colCount)
                colCount = drawRange.ColumnCount;
            if (drawRange.FirstRow < startRow)
                startRow = drawRange.FirstRow;
            if (drawRange.FirstColumn < startColumn)
                startColumn = drawRange.FirstColumn;
            
            //Create ampty ArrayList
            ArrayList tablePartList = new ArrayList();
            //split worksheet into a tables collection
            GetTablePart(excelWorksheet, doc, tablePartList, startColumn, startRow, colCount, rowCount);

            return tablePartList;
        }

        /// <summary>
        /// This Method splits worksheet into a tables collection.
        /// </summary>
        /// <param name="excelWorksheet">Input Worksheet</param>
        /// <param name="doc">Parent document</param>
        /// <param name="tablePartList">ArrayList where tables will be stored</param>
        /// <param name="columnStartIndex">Index of the column in Excel worksheet that will be the first column of Word table</param>
        /// <param name="rowStartIndex">Index of the row in Excel worksheet that will be the first row of Word table</param>
        /// <param name="columnCount">Column index of a last cell that contains data in the Excel worksheet</param>
        /// <param name="rowCount">Row index of a last cell that contains data in the Excel worksheet</param>
        private void GetTablePart(  Worksheet excelWorksheet, 
                                    Document doc, 
                                    ArrayList tablePartList, 
                                    int columnStartIndex,
                                    int rowStartIndex,
                                    int columnCount, 
                                    int rowCount)
        {
            if (columnCount != 0 && rowCount != 0)
            {
                //Calculate max width of Words table. 
                //Then we will add columns to Word table while it's width is < maxWidth
                Aspose.Words.PageSetup setup = doc.FirstSection.PageSetup;
                double maxWidth = setup.PageWidth - setup.LeftMargin - setup.RightMargin;
                int newColumnStartIndex = 0;
                double currentWidth = 0;
                for (int columnIndex = columnStartIndex; columnIndex <= columnCount; columnIndex++)
                {
                    //Calculate width of current Word table
                    currentWidth += ConvertUtil.PixelToPoint(excelWorksheet.Cells.GetColumnWidthPixel(columnIndex));
                    newColumnStartIndex = columnIndex;
                    //If width of table > maxWidth then break loop
                    if (currentWidth > maxWidth && columnIndex != columnStartIndex)
                    {
                        break;
                    }
                }
                //Create a new Word table
                Table wordsTable = new Table(doc);
                //Loop through rows in the Excel worksheet
                for (int rowIndex = rowStartIndex; rowIndex < rowCount; rowIndex++)
                {
                    //Create new row
                    Aspose.Words.Tables.Row wordsRow = new Aspose.Words.Tables.Row(doc);
                    //Get cllection of Excel cells
                    Aspose.Cells.Cells cells = excelWorksheet.Cells;
                    //Set height of current row
                    wordsRow.RowFormat.Height = ConvertUtil.PixelToPoint(cells.GetRowHeightPixel(rowIndex));
                    //Append current row to current table.
                    wordsTable.AppendChild(wordsRow);
                    //Loop through columns and add columns to Word table while table's width < maxWidth
                    for (int columnIndex = columnStartIndex; columnIndex < newColumnStartIndex; columnIndex++)
                    {
                        //Convert Excel cell to Word cell
                        Aspose.Words.Tables.Cell wordsCell = ImportExcelCell(doc, cells, rowIndex, columnIndex);
                        //Insert cell into rhe row
                        wordsRow.AppendChild(wordsCell);
                    }
                }

                // We want the table to take only as much of the page as required.
                wordsTable.PreferredWidth = PreferredWidth.Auto;

                //Add Word table to ArrayList
                tablePartList.Add(wordsTable);

                if (newColumnStartIndex < columnCount)
                {
                    //Start next table from newColumnStartIndex
                    GetTablePart(excelWorksheet, doc, tablePartList, newColumnStartIndex, rowStartIndex,  columnCount, rowCount);
                }
            }
        }

        /// <summary>
        /// Convert Excel Cell to Word Cell
        /// </summary>
        /// <param name="doc">Parent document</param>
        /// <param name="cells">Excel cells collection</param>
        /// <param name="rowIndex">Row index</param>
        /// <param name="columnIndex">Column index</param>
        /// <returns>Word Cell</returns>
        private Aspose.Words.Tables.Cell ImportExcelCell(Document doc, Aspose.Cells.Cells cells, int rowIndex, int columnIndex)
        {
            //Create a new Word Cell
            Aspose.Words.Tables.Cell wordsCell = new Aspose.Words.Tables.Cell(doc);
            //Get Excel cell from collection
            Aspose.Cells.Cell excelCell = cells[rowIndex, columnIndex];
            //Set cell width
            double cellWidth = ConvertUtil.PixelToPoint(cells.GetColumnWidthPixel(columnIndex));
            wordsCell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(cellWidth);
            wordsCell.CellFormat.Width = ConvertUtil.PixelToPoint(cellWidth);
            //Set background color
            wordsCell.CellFormat.Shading.ForegroundPatternColor = excelCell.GetDisplayStyle().ForegroundColor;
            wordsCell.CellFormat.Shading.BackgroundPatternColor = excelCell.GetDisplayStyle().BackgroundColor;
            //Set background texture
            wordsCell.CellFormat.Shading.Texture = ConvertBackgroundTexture(excelCell.GetDisplayStyle().Pattern);
            //Import borders from Excel cell to Word cell
            ImportBorders(wordsCell, excelCell);
            //Set vertical alignment
            wordsCell.CellFormat.VerticalAlignment = ConvertVerticalAlignment(excelCell.GetDisplayStyle().VerticalAlignment);
            //If Excel cells is merged then merge cells in Word Table
            wordsCell.CellFormat.VerticalMerge = ConvertVerticalCellMerge(excelCell);
            wordsCell.CellFormat.HorizontalMerge = ConvertHorizontalCellMerge(excelCell);
            //Create paragraph that will containc content of cell
            Paragraph wordsParagraph = new Paragraph(doc);
            //Set horizontal alignment
            wordsParagraph.ParagraphFormat.Alignment = ConvertHorizontalAlignment(excelCell.GetDisplayStyle().HorizontalAlignment);
            //Get text with formating from Excel cell as collection Run
            ArrayList wordRuns = GetTextFromCell(excelCell, doc);
            foreach (Run run in wordRuns)
            {
                wordsParagraph.AppendChild(run);
            }
            //Import formating of the cell
            ImportFont(wordsParagraph.ParagraphBreakFont, excelCell.GetDisplayStyle().Font);
            //Insert paragrahp with content into cell
            wordsCell.AppendChild(wordsParagraph);
            //If Excel cell contains drawing object then convert this object and insert into Word cell
            InsertDrawingObject(excelCell, wordsCell);

            return wordsCell;
        }

        /// <summary>
        /// Inserts excel drawing object into Word cell
        /// </summary>
        /// <param name="excelCell">Excel cell</param>
        /// <param name="wordsCell">Word cell</param>
        private void InsertDrawingObject(Aspose.Cells.Cell excelCell, Aspose.Words.Tables.Cell wordsCell)
        {
            //If current cell is horizontaly merged with previose cell then we should calculate offset of Shape
            double leftOffset = GetAdditionalHorizontalOffset(excelCell);
            //If current cell is verticaly merged with previose cell then we should calculate offset of Shape
            double topOffset = GetAdditionalVerticalOffset(excelCell);

            if (mPicturesCollection.ContainsKey(excelCell))
            {
                //Get Picture object from HashTable
                Aspose.Cells.Drawing.Picture excelPicture = (Aspose.Cells.Drawing.Picture)mPicturesCollection[excelCell];
                //Convert Excel Picture to Word Shape
                Aspose.Words.Drawing.Shape wordsShape = ConvertPictureToShape(excelPicture, wordsCell.Document);
                wordsShape.Left += leftOffset;
                wordsShape.Top += topOffset;
                //Insert Shape into current cell
                wordsCell.LastParagraph.AppendChild(wordsShape);
            }
            if (mChartsCollection.ContainsKey(excelCell))
            {
                //Get Chart object from HashTable
                Aspose.Cells.Charts.Chart excelChart = (Aspose.Cells.Charts.Chart)mChartsCollection[excelCell];
                //Convert Excel Chart to Word Shape
                Aspose.Words.Drawing.Shape wordsShape = ConvertCartToShape(excelChart, wordsCell.Document);
                if (wordsShape != null)
                {
                    wordsShape.Left += leftOffset;
                    wordsShape.Top += topOffset;
                    wordsCell.LastParagraph.AppendChild(wordsShape);
                }
            }
            if (mCheckBoxesCollection.ContainsKey(excelCell))
            {
                //Get CheckBox object from HashTable
                Aspose.Cells.Drawing.CheckBox excelCheckBox = (Aspose.Cells.Drawing.CheckBox)mCheckBoxesCollection[excelCell];
                //Insert CheckBox into the current cell
                InsertCheckBox(excelCheckBox, wordsCell);
            }
            if (mTextBoxesCollection.ContainsKey(excelCell))
            {
                //Get TextBox object from HashTable
                Aspose.Cells.Drawing.TextBox excelTextBox = (Aspose.Cells.Drawing.TextBox)mTextBoxesCollection[excelCell];
                //Convert Excel TextBox to Word TextBox
                Aspose.Words.Drawing.Shape textBox = ConvertTextBoxToShape(excelTextBox, wordsCell.Document);
                textBox.Left += leftOffset;
                textBox.Top += topOffset;
                //Insert Shape into current cell
                wordsCell.LastParagraph.AppendChild(textBox);
            }
            if (mShapesCollection.ContainsKey(excelCell))
            {
                //Get Shape object from HashTable
                Shape excelShape = (Shape)mShapesCollection[excelCell];
                //Convert Excel Shape to Word Shape
                Aspose.Words.Drawing.Shape wordsShape = ConvertShapeToShape(excelShape, wordsCell.Document);
                wordsShape.Left += leftOffset;
                wordsShape.Top += topOffset;
                //Insert Shape into current cell
                wordsCell.LastParagraph.AppendChild(wordsShape);
            }
        }

        /// <summary>
        /// Extract objects like Pictures, Charts, etc and store in the HashTable
        /// </summary>
        /// <param name="excelWorksheet">Excel worksheet</param>
        /// <returns>Range that contains drawing object</returns>
        private Aspose.Cells.Range ExtractDrawingObjects(Worksheet excelWorksheet)
        {
            int rowIndex = 0;
            int columnIndex = 0;
            int lastRow = 0;
            int lastColumn = 0;
            int firstRow = excelWorksheet.Cells.Rows.Count;
            int firstColumn = excelWorksheet.Cells.Columns.Count;

            //Get collection of Pictures in current worksheet
            Aspose.Cells.Drawing.PictureCollection pictures = excelWorksheet.Pictures;
            foreach (Aspose.Cells.Drawing.Picture picture in pictures)
            {
                rowIndex = picture.UpperLeftRow;
                columnIndex = picture.UpperLeftColumn;
                if (columnIndex > lastColumn)
                    lastColumn = columnIndex;
                if (rowIndex > lastRow)
                    lastRow = rowIndex;
                if (columnIndex < firstColumn)
                    firstColumn = columnIndex;
                if (rowIndex < firstRow)
                    firstRow = rowIndex;
                //Add Picture to HashTable. Key is upper left cell and value is Picture
                mPicturesCollection.Add(excelWorksheet.Cells[rowIndex, columnIndex], picture);
            }
            //Get collection of Charts in current worksheet
            Aspose.Cells.Charts.ChartCollection charts = excelWorksheet.Charts;
            foreach (Aspose.Cells.Charts.Chart chart in charts)
            {
                rowIndex = chart.ChartObject.UpperLeftRow;
                columnIndex = chart.ChartObject.UpperLeftColumn;
                if (columnIndex > lastColumn)
                    lastColumn = columnIndex;
                if (rowIndex > lastRow)
                    lastRow = rowIndex;
                if (columnIndex < firstColumn)
                    firstColumn = columnIndex;
                if (rowIndex < firstRow)
                    firstRow = rowIndex;
                //Add Chart to HashTable. Key is upper left cell and value is Chart
                mChartsCollection.Add(excelWorksheet.Cells[rowIndex, columnIndex], chart);
            }
            //Get collection of CheckBoxes in current worksheet
            Aspose.Cells.Drawing.CheckBoxCollection checkBoxes = excelWorksheet.CheckBoxes;
            foreach (Aspose.Cells.Drawing.CheckBox checkBox in checkBoxes)
            {
                rowIndex = checkBox.UpperLeftRow;
                columnIndex = checkBox.UpperLeftColumn;
                if (columnIndex > lastColumn)
                    lastColumn = columnIndex;
                if (rowIndex > lastRow)
                    lastRow = rowIndex;
                if (columnIndex < firstColumn)
                    firstColumn = columnIndex;
                if (rowIndex < firstRow)
                    firstRow = rowIndex;
                //Add CheckBox to HashTable. Key is upper left cell and value is ChekBox
                mCheckBoxesCollection.Add(excelWorksheet.Cells[rowIndex, columnIndex], checkBox);
            }
            //Get collection of TextBoxes in current worksheet
            Aspose.Cells.Drawing.TextBoxCollection textBoxes = excelWorksheet.TextBoxes;
            foreach (Aspose.Cells.Drawing.TextBox textBox in textBoxes)
            {
                rowIndex = textBox.UpperLeftRow;
                columnIndex = textBox.UpperLeftColumn;
                if (columnIndex > lastColumn)
                    lastColumn = columnIndex;
                if (rowIndex > lastRow)
                    lastRow = rowIndex;
                if (columnIndex < firstColumn)
                    firstColumn = columnIndex;
                if (rowIndex < firstRow)
                    firstRow = rowIndex;
                //Add TextBox to HashTable. Key is upper left cell and value is ChekBox
                mTextBoxesCollection.Add(excelWorksheet.Cells[rowIndex, columnIndex], textBox);
            }
            //Get collection of Shapes
            Aspose.Cells.Drawing.ShapeCollection shapes = excelWorksheet.Shapes;
            foreach (Shape shape in shapes)
            {
                if (shape.MsoDrawingType == MsoDrawingType.Line ||
                    shape.MsoDrawingType == MsoDrawingType.Arc ||
                    shape.MsoDrawingType == MsoDrawingType.Oval ||
                    shape.MsoDrawingType == MsoDrawingType.Rectangle)
                {
                    rowIndex = shape.UpperLeftRow;
                    columnIndex = shape.UpperLeftColumn;
                    if (columnIndex > lastColumn)
                        lastColumn = columnIndex;
                    if (rowIndex > lastRow)
                        lastRow = rowIndex;
                    if (columnIndex < firstColumn)
                        firstColumn = columnIndex;
                    if (rowIndex < firstRow)
                        firstRow = rowIndex;
                    //Add TextBox to HashTable. Key is upper left cell and value is ChekBox
                    mShapesCollection.Add(excelWorksheet.Cells[rowIndex, columnIndex], shape);
                }
            }

            Aspose.Cells.Range range = excelWorksheet.Cells.CreateRange(firstRow, firstColumn, lastRow + 1, lastColumn + 1);

            return range;
        }

        /// <summary>
        /// Import boreders from Excel cell to Word Cell
        /// </summary>
        /// <param name="wordsCell">Destination Word cell</param>
        /// <param name="excelCell">Source Excel cell</param>
        private void ImportBorders(Aspose.Words.Tables.Cell wordsCell, Aspose.Cells.Cell excelCell)
        {
            //Set line style
            wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Bottom].LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.BottomBorder].LineStyle);
            //Set line color
            if (wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Bottom].LineStyle != LineStyle.None)
                wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Bottom].Color = excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.BottomBorder].Color;

            wordsCell.CellFormat.Borders[Aspose.Words.BorderType.DiagonalDown].LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.DiagonalDown].LineStyle);
            if (wordsCell.CellFormat.Borders[Aspose.Words.BorderType.DiagonalDown].LineStyle!= LineStyle.None)
                wordsCell.CellFormat.Borders[Aspose.Words.BorderType.DiagonalDown].Color = excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.DiagonalDown].Color;

            wordsCell.CellFormat.Borders[Aspose.Words.BorderType.DiagonalUp].LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.DiagonalUp].LineStyle);
            if (wordsCell.CellFormat.Borders[Aspose.Words.BorderType.DiagonalUp].LineStyle!= LineStyle.None)
                wordsCell.CellFormat.Borders[Aspose.Words.BorderType.DiagonalUp].Color = excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.DiagonalUp].Color;

            wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Left].LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.LeftBorder].LineStyle);
            if (wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Left].LineStyle != LineStyle.None)
                wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Left].Color = excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.LeftBorder].Color;

            wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Right].LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.RightBorder].LineStyle);
            if (wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Right].LineStyle != LineStyle.None)
                wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Right].Color = excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.RightBorder].Color;

            wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Top].LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.TopBorder].LineStyle);
            if (wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Top].LineStyle != LineStyle.None)
                wordsCell.CellFormat.Borders[Aspose.Words.BorderType.Top].Color = excelCell.GetDisplayStyle().Borders[Aspose.Cells.BorderType.TopBorder].Color;
        }

        /// <summary>
        /// Convert Excel BorderType to Word LineStyle
        /// </summary>
        /// <param name="lineType">Excel LineType</param>
        /// <returns>Word LineStyle</returns>
        private LineStyle ConvertLineStyle(CellBorderType lineType)
        {
            LineStyle wordsLineStyle = LineStyle.None;
            switch (lineType)
            {
                case CellBorderType.DashDot:
                    {
                        wordsLineStyle = LineStyle.DotDash;
                        break;
                    }
                case CellBorderType.DashDotDot:
                    {
                        wordsLineStyle = LineStyle.DotDotDash;
                        break;
                    }
                case CellBorderType.Dashed:
                    {
                        wordsLineStyle = LineStyle.DashSmallGap;
                        break;
                    }
                case CellBorderType.Dotted:
                    {
                        wordsLineStyle = LineStyle.Dot;
                        break;
                    }
                case CellBorderType.Double:
                    {
                        wordsLineStyle = LineStyle.Double;
                        break;
                    }
                case CellBorderType.Hair:
                    {
                        wordsLineStyle = LineStyle.Hairline;
                        break;
                    }
                case CellBorderType.Medium:
                    {
                        wordsLineStyle = LineStyle.Single;
                        break;
                    }
                case CellBorderType.MediumDashDot:
                    {
                        wordsLineStyle = LineStyle.DotDash;
                        break;
                    }
                case CellBorderType.MediumDashDotDot:
                    {
                        wordsLineStyle = LineStyle.DotDotDash;
                        break;
                    }
                case CellBorderType.MediumDashed:
                    {
                        wordsLineStyle = LineStyle.DashSmallGap;
                        break;
                    }
                case CellBorderType.None:
                    {
                        wordsLineStyle = LineStyle.None;
                        break;
                    }
                case CellBorderType.SlantedDashDot:
                    {
                        wordsLineStyle = LineStyle.DotDash;
                        break;
                    }
                case CellBorderType.Thick:
                    {
                        wordsLineStyle = LineStyle.Thick;
                        break;
                    }
                case CellBorderType.Thin:
                    {
                        wordsLineStyle = LineStyle.Single;
                        break;
                    }
                default:
                    {
                        wordsLineStyle = LineStyle.None;
                        break;
                    }
            }

            return wordsLineStyle;
        }

        /// <summary>
        /// Conver Excel Undeline to Word Underline
        /// </summary>
        /// <param name="underlineType">Excel UnderlineType</param>
        /// <returns>Word Underline</returns>
        private Underline ConvertUnderline(FontUnderlineType underlineType)
        {
            Underline wordsUnderline = Underline.None;

            switch (underlineType)
            {
                case FontUnderlineType.Accounting:
                    {
                        wordsUnderline = Underline.Wavy;
                        break;
                    }
                case FontUnderlineType.Double:
                    {
                        wordsUnderline = Underline.Double;
                        break;
                    }
                case FontUnderlineType.DoubleAccounting:
                    {
                        wordsUnderline = Underline.WavyDouble;
                        break;
                    }
                case FontUnderlineType.None:
                    {
                        wordsUnderline = Underline.None;
                        break;
                    }
                case FontUnderlineType.Single:
                    {
                        wordsUnderline = Underline.Single;
                        break;
                    }
                default:
                    {
                        wordsUnderline = Underline.None;
                        break;
                    }
            }

            return wordsUnderline;
        }

        /// <summary>
        /// Conver Excel VerticalAlignment to Word CellVerticalAlignment
        /// </summary>
        /// <param name="alignmentType">Excel AlignmentType</param>
        /// <returns>Word CellVerticalAlignment</returns>
        private CellVerticalAlignment ConvertVerticalAlignment(TextAlignmentType alignmentType)
        {
            CellVerticalAlignment wordsAlignment = CellVerticalAlignment.Top;

            switch (alignmentType)
            {
                case TextAlignmentType.Bottom:
                    {
                        wordsAlignment = CellVerticalAlignment.Bottom;
                        break;
                    }
                case TextAlignmentType.Center:
                    {
                        wordsAlignment = CellVerticalAlignment.Center;
                        break;
                    }
                case TextAlignmentType.Top:
                    {
                        wordsAlignment = CellVerticalAlignment.Top;
                        break;
                    }
                default:
                    {
                        wordsAlignment = CellVerticalAlignment.Top;
                        break;
                    }
            }

            return wordsAlignment;
        }

        /// <summary>
        /// Conver Excel HorizontalAlignment to Word ParagraphAlignment
        /// </summary>
        /// <param name="alignmentType">Excel AlignmentType</param>
        /// <returns>Word ParagraphAlignment</returns>
        private ParagraphAlignment ConvertHorizontalAlignment(TextAlignmentType alignmentType)
        {
            ParagraphAlignment wordsAlignment = ParagraphAlignment.Left;

            switch (alignmentType)
            {
                case TextAlignmentType.Center:
                    {
                        wordsAlignment = ParagraphAlignment.Center;
                        break;
                    }
                case TextAlignmentType.Distributed:
                    {
                        wordsAlignment = ParagraphAlignment.Distributed;
                        break;
                    }
                case TextAlignmentType.Justify:
                    {
                        wordsAlignment = ParagraphAlignment.Justify;
                        break;
                    }
                case TextAlignmentType.Left:
                    {
                        wordsAlignment = ParagraphAlignment.Left;
                        break;
                    }
                case TextAlignmentType.Right:
                    {
                        wordsAlignment = ParagraphAlignment.Right;
                        break;
                    }
                default:
                    {
                        wordsAlignment = ParagraphAlignment.Left;
                        break;
                    }
            }

            return wordsAlignment;
        }

        /// <summary>
        /// Convert Excel HorizontalCellMerge to Word HorizontalCellMerge
        /// </summary>
        /// <param name="excelCell">Input Excel cell</param>
        /// <returns>CellMerge type</returns>
        private CellMerge ConvertHorizontalCellMerge(Aspose.Cells.Cell excelCell)
        {
            //By default cells are not merged
            CellMerge wordsCellMerge = CellMerge.None;
            //Get merged region
            Aspose.Cells.Range mergedRange = excelCell.GetMergedRange();
            if (mergedRange == null)
            {
                //Cells are not merged
                wordsCellMerge = CellMerge.None;
            }
            else
            {
                if (excelCell.Column == mergedRange.FirstColumn && mergedRange.ColumnCount > 1)
                {
                    //Cell is merged with next
                    wordsCellMerge = CellMerge.First;
                }
                else if (mergedRange.ColumnCount > 1)
                {
                    //Cell is merged with previouse
                    wordsCellMerge = CellMerge.Previous;
                }
                else
                {
                    //Cell is not merged
                    wordsCellMerge = CellMerge.None;
                }
            }
            return wordsCellMerge;
        }

        /// <summary>
        /// Convert Excel VerticalCellMerge to Word VerticalCellMerge
        /// </summary>
        /// <param name="excelCell">Input Excel cell</param>
        /// <returns>CellMerge type</returns>
        private CellMerge ConvertVerticalCellMerge(Aspose.Cells.Cell excelCell)
        {
            //By default cells are not merged
            CellMerge wordsCellMerge = CellMerge.None;
            //Get merged region
            Aspose.Cells.Range mergedRange = excelCell.GetMergedRange();

            if (mergedRange == null)
            {
                //Cells are not merged
                wordsCellMerge = CellMerge.None;
            }
            else
            {
                if ((excelCell.Row.Equals(mergedRange.FirstRow)) && (mergedRange.RowCount > 1))
                {
                    //Cell is merged with next
                    wordsCellMerge = CellMerge.First;
                }
                else if (mergedRange.RowCount > 1)
                {
                    //Cell is merged with previouse
                    wordsCellMerge = CellMerge.Previous;
                }
                else
                {
                    //Cell is not merged
                    wordsCellMerge = CellMerge.None;
                }
            }
            return wordsCellMerge;
        }

        /// <summary>
        /// Convert Excel MsoLineStyle to Word ShapeLineStyle
        /// </summary>
        /// <param name="lineStyle">Excel MsoLineStyle</param>
        /// <returns>Word ShapeLineStyle</returns>
        private Aspose.Words.Drawing.ShapeLineStyle ConvertDrawingLineStyle(MsoLineStyle lineStyle)
        {
            Aspose.Words.Drawing.ShapeLineStyle wordsLineStyle = Aspose.Words.Drawing.ShapeLineStyle.Single;

            switch (lineStyle)
            {
                case MsoLineStyle.Single:
                    {
                        wordsLineStyle = Aspose.Words.Drawing.ShapeLineStyle.Single;
                        break;
                    }
                case MsoLineStyle.ThickBetweenThin:
                    {
                        wordsLineStyle = Aspose.Words.Drawing.ShapeLineStyle.Triple;
                        break;
                    }
                case MsoLineStyle.ThickThin:
                    {
                        wordsLineStyle = Aspose.Words.Drawing.ShapeLineStyle.ThickThin;
                        break;
                    }
                case MsoLineStyle.ThinThick:
                    {
                        wordsLineStyle = Aspose.Words.Drawing.ShapeLineStyle.ThinThick;
                        break;
                    }
                case MsoLineStyle.ThinThin:
                    {
                        wordsLineStyle = Aspose.Words.Drawing.ShapeLineStyle.Double;
                        break;
                    }
            }

            return wordsLineStyle;
        }

        /// <summary>
        /// Convert Excel MsoLineDashStyle to Word DashStyle
        /// </summary>
        /// <param name="dashStyle">Excel MsoLineDashStyle</param>
        /// <returns>Word DashStyle</returns>
        private Aspose.Words.Drawing.DashStyle ConvertDrawingDashStyle(MsoLineDashStyle dashStyle)
        {
            Aspose.Words.Drawing.DashStyle wordsDashStyle = Aspose.Words.Drawing.DashStyle.Solid;
            switch (dashStyle)
            {
                case MsoLineDashStyle.Dash:
                    {
                        wordsDashStyle = Aspose.Words.Drawing.DashStyle.Dash;
                        break;
                    }
                case MsoLineDashStyle.DashDot:
                    {
                        wordsDashStyle = Aspose.Words.Drawing.DashStyle.DashDot;
                        break;
                    }
                case MsoLineDashStyle.DashDotDot:
                    {
                        wordsDashStyle = Aspose.Words.Drawing.DashStyle.LongDashDotDot;
                        break;
                    }
                case MsoLineDashStyle.DashLongDash:
                    {
                        wordsDashStyle = Aspose.Words.Drawing.DashStyle.LongDash;
                        break;
                    }
                case MsoLineDashStyle.DashLongDashDot:
                    {
                        wordsDashStyle = Aspose.Words.Drawing.DashStyle.LongDashDot;
                        break;
                    }
                case MsoLineDashStyle.RoundDot:
                    {
                        wordsDashStyle = Aspose.Words.Drawing.DashStyle.Dot;
                        break;
                    }
                case MsoLineDashStyle.Solid:
                    {
                        wordsDashStyle = Aspose.Words.Drawing.DashStyle.Solid;
                        break;
                    }
                case MsoLineDashStyle.SquareDot:
                    {
                        wordsDashStyle = Aspose.Words.Drawing.DashStyle.ShortDot;
                        break;
                    }
            }
            return wordsDashStyle;
        }

        /// <summary>
        /// Convert Excel TextOrientationType to Word LayoutFlow
        /// </summary>
        /// <param name="textOrientation">Excel TextOrientationType</param>
        /// <returns>Word LayoutFlow</returns>
        private Aspose.Words.Drawing.LayoutFlow ConvertDrawingTextOrientationType(TextOrientationType textOrientation)
        {
            Aspose.Words.Drawing.LayoutFlow wordLayoutFlow = Aspose.Words.Drawing.LayoutFlow.Horizontal;
            switch (textOrientation)
            {
                case TextOrientationType.ClockWise:
                    {
                        wordLayoutFlow = Aspose.Words.Drawing.LayoutFlow.BottomToTop;
                        break;
                    }
                case TextOrientationType.CounterClockWise:
                    {
                        wordLayoutFlow = Aspose.Words.Drawing.LayoutFlow.TopToBottom;
                        break;
                    }
                case TextOrientationType.NoRotation:
                    {
                        wordLayoutFlow = Aspose.Words.Drawing.LayoutFlow.Horizontal;
                        break;
                    }
                case TextOrientationType.TopToBottom:
                    {
                        wordLayoutFlow = Aspose.Words.Drawing.LayoutFlow.TopToBottom;
                        break;
                    }
            }
            return wordLayoutFlow;
        }

        /// <summary>
        /// Convert Excel MsoDrawingType to Word ShapeType
        /// </summary>
        /// <param name="excelShapeType">Excel MsoDrawingType</param>
        /// <returns>Word ShapeType</returns>
        private Aspose.Words.Drawing.ShapeType ConvertDrawingShapetype(MsoDrawingType excelShapeType)
        {
            Aspose.Words.Drawing.ShapeType wordsShapeType = Aspose.Words.Drawing.ShapeType.Line;
            switch (excelShapeType)
            {
                case MsoDrawingType.Arc:
                    {
                        wordsShapeType = Aspose.Words.Drawing.ShapeType.Arc;
                        break;
                    }
                case MsoDrawingType.Line:
                    {
                        wordsShapeType = Aspose.Words.Drawing.ShapeType.Line;
                        break;
                    }
                case MsoDrawingType.Oval:
                    {
                        wordsShapeType = Aspose.Words.Drawing.ShapeType.Ellipse;
                        break;
                    }
                case MsoDrawingType.Rectangle:
                    {
                        wordsShapeType = Aspose.Words.Drawing.ShapeType.Rectangle;
                        break;
                    }
            }
            return wordsShapeType;
        }

        /// <summary>
        /// Convert excel BackgroundType to Word TextureIndex
        /// </summary>
        /// <param name="excelTextureType">excel BackgroundType</param>
        /// <returns>Word TextureIndex</returns>
        private Aspose.Words.TextureIndex ConvertBackgroundTexture(Aspose.Cells.BackgroundType excelTextureType)
        {
            Aspose.Words.TextureIndex wordsTextureIndex = Aspose.Words.TextureIndex.TextureNone;
            switch (excelTextureType)
            {
                case BackgroundType.DiagonalCrosshatch:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalCross;
                        break;
                    }
                case BackgroundType.DiagonalStripe:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalDown;
                        break;
                    }
                case BackgroundType.Gray12:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.Texture12Pt5Percent;
                        break;
                    }
                case BackgroundType.Gray25:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.Texture25Percent;
                        break;
                    }
                case BackgroundType.Gray50:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.Texture50Percent;
                        break;
                    }
                case BackgroundType.Gray6:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.Texture10Percent;
                        break;
                    }
                case BackgroundType.Gray75:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.Texture75Percent;
                        break;
                    }
                case BackgroundType.HorizontalStripe:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureHorizontal;
                        break;
                    }
                case BackgroundType.None:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureNone;
                        break;
                    }
                case BackgroundType.ReverseDiagonalStripe:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalUp;
                        break;
                    }
                case BackgroundType.Solid:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureSolid;
                        break;
                    }
                case BackgroundType.ThickDiagonalCrosshatch:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureDarkDiagonalCross;
                        break;
                    }
                case BackgroundType.ThinDiagonalCrosshatch:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalCross;
                        break;
                    }
                case BackgroundType.ThinDiagonalStripe:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalCross;
                        break;
                    }
                case BackgroundType.ThinHorizontalCrosshatch:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureHorizontal;
                        break;
                    }
                case BackgroundType.ThinHorizontalStripe:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureHorizontal;
                        break;
                    }
                case BackgroundType.ThinReverseDiagonalStripe:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalUp;
                        break;
                    }
                case BackgroundType.ThinVerticalStripe:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureVertical;
                        break;
                    }
                case BackgroundType.VerticalStripe:
                    {
                        wordsTextureIndex = Aspose.Words.TextureIndex.TextureVertical;
                        break;
                    }

            }
            return wordsTextureIndex;
        }

        /// <summary>
        /// Get text with formating from Excel cell
        /// </summary>
        /// <param name="cell">Input Excel cell</param>
        /// <param name="doc">Parent document</param>
        /// <returns>Run that contains text with formating</returns>
        private ArrayList GetTextFromCell(Aspose.Cells.Cell cell, Document doc)
        {
            //Create new array list 
            //We will store runs in this list
            ArrayList wordRuns = new ArrayList();

            //Get Chracters objects from the cell
            FontSetting[] charactersList = cell.GetCharacters();
            if (charactersList == null)
                charactersList = new FontSetting[0];
          
            //Get sring value from the cell
            string cellValue = cell.StringValue;

            //If there is some formatig in the cell 
            //charactersList will contains one or more Characters objects
            //And we should create collection of Runs that will represent this formating
            if (charactersList.Length > 0)
            {
                int startIndex = ((FontSetting)charactersList[0]).StartIndex;
                if (startIndex > 0)
                {
                    //Create first run
                    Run firstRun = new Run(doc);
                    //Set text of first run. That will be substring that starts from 0
                    //And ens and the start position of first Characters object

                    //Sometimes Aspose.Cell returns incorrect index so we should check it
                    if (startIndex < cellValue.Length)
                        firstRun.Text = cellValue.Substring(0, startIndex);
                    else
                        firstRun.Text = cellValue.Substring(0);
                    //Formation of first run will be the same as formating of whole cell
                    ImportFont(firstRun.Font, cell.GetDisplayStyle().Font);
                    wordRuns.Add(firstRun);
                }

                //Loop through all Character objects
                foreach (FontSetting chars in charactersList)
                {
                    Run run = new Run(doc);
                    //We should check index to avoid errors
                    if (chars.StartIndex < cellValue.Length)
                    {
                        if (cellValue.Length > (chars.StartIndex + chars.Length) && chars.Length > 0)
                            run.Text = cellValue.Substring(chars.StartIndex, chars.Length);
                        else
                            run.Text = cellValue.Substring(chars.StartIndex);
                    }
                    //Convert Excel Font to Words Font
                    ImportFont(run.Font, chars.Font);
                    wordRuns.Add(run);
                }
            }
            //Othervice there will be only one run 
            else
            {
                Run run = new Run(doc);
                run.Text = cellValue;
                //Convert Excel Font to Words Font
                ImportFont(run.Font, cell.GetDisplayStyle().Font);
                wordRuns.Add(run);
            }

            return wordRuns;
        }

        /// <summary>
        /// Convert Excel Font to Word Font
        /// </summary>
        /// <param name="wordsFont">Word Font</param>
        /// <param name="excelFont">Excel Font</param>
        private void ImportFont(Aspose.Words.Font wordsFont, Aspose.Cells.Font excelFont)
        {
            wordsFont.Name = excelFont.Name;
            wordsFont.Bold = excelFont.IsBold;
            wordsFont.Color = excelFont.Color;
            wordsFont.Italic = excelFont.IsItalic;
            wordsFont.StrikeThrough = excelFont.IsStrikeout;
            wordsFont.Subscript = excelFont.IsSubscript;
            wordsFont.Superscript = excelFont.IsSuperscript;
            wordsFont.Size = excelFont.Size;
            wordsFont.Underline = ConvertUnderline(excelFont.Underline);
        }

        /// <summary>
        /// Convert Excel PaperSize to Word PaperSize
        /// </summary>
        /// <param name="excelPaperSize">Excel PaperSize</param>
        /// <returns>Word Paper size</returns>
        private PaperSize ConvertPaperSize(PaperSizeType excelPaperSize)
        {
            PaperSize paperSize = PaperSize.A4;

            switch (excelPaperSize)
            {
                case PaperSizeType.PaperA4:
                    {
                        paperSize = PaperSize.A4;
                        break;
                    }
                case PaperSizeType.PaperA3:
                    {
                        paperSize = PaperSize.A3;
                        break;
                    }
                case PaperSizeType.PaperA5:
                    {
                        paperSize = PaperSize.A5;
                        break;
                    }
                case PaperSizeType.PaperB4:
                    {
                        paperSize = PaperSize.B4;
                        break;
                    }
                case PaperSizeType.PaperB5:
                    {
                        paperSize = PaperSize.B5;
                        break;
                    }
                case PaperSizeType.Paper10x14:
                    {
                        paperSize = PaperSize.Paper10x14;
                        break;
                    }
                case PaperSizeType.Paper11x17:
                    {
                        paperSize = PaperSize.Paper11x17;
                        break;
                    }
                case PaperSizeType.PaperEnvelopeDL:
                    {
                        paperSize = PaperSize.EnvelopeDL;
                        break;
                    }
                case PaperSizeType.PaperExecutive:
                    {
                        paperSize = PaperSize.Executive;
                        break;
                    }
                case PaperSizeType.PaperFolio:
                    {
                        paperSize = PaperSize.Folio;
                        break;
                    }
                case PaperSizeType.PaperLedger:
                    {
                        paperSize = PaperSize.Ledger;
                        break;
                    }
                case PaperSizeType.PaperLegal:
                    {
                        paperSize = PaperSize.Legal;
                        break;
                    }
                case PaperSizeType.PaperLetter:
                    {
                        paperSize = PaperSize.Letter;
                        break;
                    }
                case PaperSizeType.PaperQuarto:
                    {
                        paperSize = PaperSize.Quarto;
                        break;
                    }
                case PaperSizeType.PaperStatement:
                    {
                        paperSize = PaperSize.Statement;
                        break;
                    }
                case PaperSizeType.PaperTabloid:
                    {
                        paperSize = PaperSize.Tabloid;
                        break;
                    }
                default:
                    {
                        paperSize = PaperSize.Letter;
                        break;
                    }
            }

            return paperSize;
        }

        /// <summary>
        /// Convert Excel PageOrientaton to Word PageOrientaton
        /// </summary>
        /// <param name="excelPageOrientation">Excel PageOrientation (Portrait or Landscape)</param>
        /// <returns>Portrait or Landscape, by default returns Portrait</returns>
        private Orientation ConvertPageOrientation(PageOrientationType excelPageOrientation)
        {
            Orientation pageOrientation = Orientation.Portrait;

            switch (excelPageOrientation)
            {
                case PageOrientationType.Portrait:
                    {
                        pageOrientation = Orientation.Portrait;
                        break;
                    }
                case PageOrientationType.Landscape:
                    {
                        pageOrientation = Orientation.Landscape;
                        break;
                    }
                default:
                    {
                        pageOrientation = Orientation.Portrait;
                        break;
                    }
            }

            return pageOrientation;
        }

        /// <summary>
        /// Calculate additional offset if excel cell is merged horizontally
        /// </summary>
        /// <param name="excelCell">Excel cell</param>
        /// <returns></returns>
        private double GetAdditionalHorizontalOffset(Aspose.Cells.Cell excelCell)
        {
            double leftOffset = 0;
            //Get merged region of excel Cell
            Aspose.Cells.Range mergedRange = excelCell.GetMergedRange();
            if (mergedRange != null)
            {
                if (excelCell.Column != mergedRange.FirstColumn && mergedRange.ColumnCount > 1)
                {
                    //Cell is merged with previouse
                    for (int columnIndex = mergedRange.FirstColumn; columnIndex < excelCell.Column; columnIndex++)
                    {
                        leftOffset += ConvertUtil.PixelToPoint(mergedRange.Worksheet.Cells.GetColumnWidthPixel(columnIndex));
                    }
                }
            }
            return leftOffset;
        }

        /// <summary>
        /// Calculate additional offset if excel cell is merged vertically
        /// </summary>
        /// <param name="excelCell">Excel cell</param>
        /// <returns></returns>
        private double GetAdditionalVerticalOffset(Aspose.Cells.Cell excelCell)
        {
            double topOffset = 0;
            //Get merged region of excel Cell
            Aspose.Cells.Range mergedRange = excelCell.GetMergedRange();
            if (mergedRange != null)
            {
                if ((!excelCell.Row.Equals(mergedRange.FirstRow)) && (mergedRange.RowCount > 1))
                {
                    //Cell is merged with previouse
                    for (int rowIndex = mergedRange.FirstRow; rowIndex < excelCell.Row; rowIndex++)
                    {
                        topOffset += ConvertUtil.PixelToPoint(mergedRange.Worksheet.Cells.GetRowHeightPixel(rowIndex));
                    }
                }
            }
            return topOffset;
        }

        /// <summary>
        /// Convert Excel Picture to Word Shape
        /// </summary>
        /// <param name="excelPicture">Excel Picture</param>
        /// <param name="doc">Parent document</param>
        /// <returns>Word Shape</returns>
        private Aspose.Words.Drawing.Shape ConvertPictureToShape(Aspose.Cells.Drawing.Picture excelPicture, DocumentBase doc)
        {
            //Create new Shape
            Aspose.Words.Drawing.Shape wordsShape = new Aspose.Words.Drawing.Shape(doc, Aspose.Words.Drawing.ShapeType.Image);
            //Set image
            wordsShape.ImageData.ImageBytes = excelPicture.Data;
            //Import Picture properties inhereted from Shape
            ImportShapeProperties(wordsShape, (Aspose.Cells.Drawing.Shape)excelPicture);
            return wordsShape;
        }

        /// <summary>
        /// Convert Excel Chart to Word Shape
        /// </summary>
        /// <param name="excelChart">Excel Chart</param>
        /// <param name="doc">Parent document</param>
        /// <returns>Word Shape</returns>
        private Aspose.Words.Drawing.Shape ConvertCartToShape(Aspose.Cells.Charts.Chart excelChart, DocumentBase doc)
        {
            //Create a new Shape
            Aspose.Words.Drawing.Shape wordsShape = new Aspose.Words.Drawing.Shape(doc, Aspose.Words.Drawing.ShapeType.Image);
            //Convert Chart to Bitmap. Now only supports to convert 2D chart to image. If the chart is 3D chart,return null. 
            Bitmap chartPicture = excelChart.ToImage();
            if (chartPicture != null)
            {
                wordsShape.ImageData.SetImage(chartPicture);
                //Import Chart properties inhereted from Shape
                ImportShapeProperties(wordsShape, (Shape)excelChart.ChartObject);
                return wordsShape;
            }
            else
            {
                return null;
            }
        }
        
        /// <summary>
        /// Insert CheckBox into a Word cell
        /// </summary>
        /// <param name="excelCheckbox">Excel CheckBox</param>
        /// <param name="parentCell">Parent Word cell</param>
        private void InsertCheckBox(Aspose.Cells.Drawing.CheckBox excelCheckbox, Aspose.Words.Tables.Cell parentCell)
        {
            //Create new temporary document
            Document doc = new Document();
            //Create instance of DocumentBuilder
            DocumentBuilder builder = new DocumentBuilder(doc);
            //Calculate size of CheckBox
            int size = (int)(ConvertUtil.PixelToPoint(excelCheckbox.Height));

            switch (excelCheckbox.CheckedValue)
            {
                case CheckValueType.Checked:
                    {
                        builder.InsertCheckBox(excelCheckbox.Name, true, size);
                        break;
                    }
                case CheckValueType.UnChecked:
                    {
                        builder.InsertCheckBox(excelCheckbox.Name, false, size);
                        break;
                    }
                default:
                    {
                        builder.InsertCheckBox(excelCheckbox.Name, false, size);
                        break;
                    }
            }
            //Write text of Excel CheckBox
            builder.Write(excelCheckbox.Text);

            //Import all content of temporary document into a destination cell
            foreach (Node node in builder.CurrentParagraph.ChildNodes)
            {
                parentCell.LastParagraph.AppendChild(parentCell.Document.ImportNode(node, true));
            }
        }

        /// <summary>
        /// Convert Excel TextBox to Word TextBox
        /// </summary>
        /// <param name="excelTextBox">Excel TextBox</param>
        /// <param name="doc">Parent document</param>
        /// <returns>Word Shape</returns>
        private Aspose.Words.Drawing.Shape ConvertTextBoxToShape(Aspose.Cells.Drawing.TextBox excelTextBox, DocumentBase doc)
        {
            //Create a new TextBox
            Aspose.Words.Drawing.Shape wordsShape = new Aspose.Words.Drawing.Shape(doc, Aspose.Words.Drawing.ShapeType.TextBox);
            //Import TextBox properties inhereted from Shape
            ImportShapeProperties(wordsShape, (Shape)excelTextBox);
            //Import TextBox properties
            wordsShape.TextBox.LayoutFlow = ConvertDrawingTextOrientationType(excelTextBox.TextOrientationType);
            //Import text
            Run run = new Run(doc);
            if (!string.IsNullOrEmpty(excelTextBox.Text))
                run.Text = excelTextBox.Text;
            else
                run.Text = string.Empty;
            //Import text formating
            ImportFont(run.Font, excelTextBox.Font);
            //Create paragraph
            Paragraph paragraph = new Paragraph(doc);
            //Import horizontal alignment
            paragraph.ParagraphFormat.Alignment = ConvertHorizontalAlignment(excelTextBox.TextHorizontalAlignment);
            //Insert text into the paragraph
            paragraph.AppendChild(run);
            //insert Pragraph into textbox
            wordsShape.AppendChild(paragraph);
            return wordsShape;
        }

        /// <summary>
        /// Convert Excel Shape to Word Shape
        /// </summary>
        /// <param name="excelShape">Excel Shape</param>
        /// <param name="doc">Parent document</param>
        /// <returns>Word Shape</returns>
        private Aspose.Words.Drawing.Shape ConvertShapeToShape(Aspose.Cells.Drawing.Shape excelShape, DocumentBase doc)
        {
            //Create words Shape
            Aspose.Words.Drawing.Shape wordsShape = new Aspose.Words.Drawing.Shape(doc, ConvertDrawingShapetype(excelShape.MsoDrawingType));
            //Import properties
            ImportShapeProperties(wordsShape, excelShape);

            wordsShape.Stroked = true;
            wordsShape.Filled = true;

            return wordsShape;
        }

        /// <summary>
        /// Import properties of Excel shape
        /// </summary>
        /// <param name="wordsShape">Word Shape</param>
        /// <param name="excelShape">Excel Shape</param>
        private void ImportShapeProperties(Aspose.Words.Drawing.Shape wordsShape, Aspose.Cells.Drawing.Shape excelShape)
        {
            //Import size of TextBox
            wordsShape.Height = ConvertUtil.PixelToPoint(excelShape.Height);//1pt=1px*0.75
            wordsShape.Width = ConvertUtil.PixelToPoint(excelShape.Width);
            //Import horizontal offset
            wordsShape.Left = ConvertUtil.PixelToPoint(excelShape.Left);
            //Import vertical offset
            wordsShape.Top = ConvertUtil.PixelToPoint(excelShape.Top);

            //Import Filling
            if (excelShape.FillFormat.IsVisible)
            {
                wordsShape.Filled = true;
                wordsShape.Fill.Color = excelShape.FillFormat.ForeColor;
            }
            else
            {
                wordsShape.Filled = false;
            }
            //Import LineFormat (borders)
            if (excelShape.LineFormat.IsVisible)
            {
                wordsShape.Stroked = true;
                //Set LineStyle
                wordsShape.Stroke.LineStyle = ConvertDrawingLineStyle(excelShape.LineFormat.Style);
                //Set DashStyle
                wordsShape.Stroke.DashStyle = ConvertDrawingDashStyle(excelShape.LineFormat.DashStyle);
                //Set Weight
                wordsShape.Stroke.Weight = excelShape.LineFormat.Weight;
                //Set collors
                wordsShape.Stroke.Color = excelShape.LineFormat.ForeColor.IsEmpty ? Color.Black : excelShape.LineFormat.ForeColor;
                wordsShape.Stroke.Color2 = excelShape.LineFormat.BackColor;
            }
            else
            {
                wordsShape.Stroked = false;
            }
            //Import link
            if (excelShape.Hyperlink != null)
            {
                wordsShape.HRef = excelShape.Hyperlink.Address;
            }
            //Import rotation
            wordsShape.Rotation = excelShape.RotationAngle;
        }


        #region Private variables

        //Create HashTables. We will store in these tables objects like Pictures, Charts, etc.
        private Hashtable mShapesCollection = new Hashtable();
        private Hashtable mPicturesCollection = new Hashtable();
        private Hashtable mChartsCollection = new Hashtable();
        private Hashtable mTextBoxesCollection = new Hashtable();
        private Hashtable mCheckBoxesCollection = new Hashtable();

        #endregion
    }
}
