// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class InsertPicture: TestUtil
    {
        [Test]
        public static void InsertImageOpenXml()
        {
            //ExStart:InsertImageOpenXml
            //GistDesc:Insert image using C#
            using WordprocessingDocument wordDocument = WordprocessingDocument.Create(ArtifactsDir + "Insert image - OpenXML.docx", WordprocessingDocumentType.Document);

            // Add the main document part
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = new Body();

            // Add a paragraph to the document
            Paragraph paragraph = new Paragraph();
            Run run = new Run();

            // Add the image to the document
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream stream = new FileStream(MyDir + "Aspose.Words.png", FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            // Generate a unique relationship ID for the image
            string imageId = mainPart.GetIdOfPart(imagePart);

            // Define the image's dimensions (width and height in pixels)
            const int emuPerPixel = 9525;
            int widthInPixels = 300;
            int heightInPixels = 200;

            // Add the image to the run
            run.AppendChild(new Drawing(
                new DW.Inline(
                    new DW.Extent()
                    {
                        Cx = widthInPixels * emuPerPixel,
                        Cy = heightInPixels * emuPerPixel
                    },
                    new DW.DocProperties()
                    {
                        Id = 1,
                        Name = "Picture 1"
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(),
                    new Graphic(
                        new GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties()
                                    {
                                        Id = 0,
                                        Name = "Image"
                                    },
                                    new PIC.NonVisualPictureDrawingProperties()
                                ),
                                new PIC.BlipFill(
                                    new Blip()
                                    {
                                        Embed = imageId,
                                        CompressionState = BlipCompressionValues.Print
                                    },
                                    new Stretch(
                                        new FillRectangle()
                                    )
                                ),
                                new PIC.ShapeProperties(
                                    new Transform2D(
                                        new Offset()
                                        {
                                            X = 0,
                                            Y = 0
                                        },
                                        new Extents()
                                        {
                                            Cx = widthInPixels * emuPerPixel,
                                            Cy = heightInPixels * emuPerPixel
                                        }
                                    ),
                                    new PresetGeometry(
                                        new AdjustValueList()
                                    )
                                    {
                                        Preset = ShapeTypeValues.Rectangle
                                    }
                                )
                            )
                        )
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }
                    )
                )
                {
                    DistanceFromTop = 0,
                    DistanceFromBottom = 0,
                    DistanceFromLeft = 0,
                    DistanceFromRight = 0,
                    EditId = "50D07946"
                }
            ));

            // Add the run to the paragraph
            paragraph.AppendChild(run);

            // Add the paragraph to the body
            body.AppendChild(paragraph);

            // Add the body to the document
            mainPart.Document.AppendChild(body);

            // Save the document
            mainPart.Document.Save();
            //ExEnd:InsertImageOpenXml
        }
    }
}
