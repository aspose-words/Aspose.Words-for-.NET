// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;

using Aspose.Words;
using Aspose.Words.Saving;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExSwfSaveOptions : ApiExampleBase
    {
        [Test]
        public void UseCustomToolTips()
        {
            Document doc = new Document(MyDir + "Document.doc");

            //ExStart
            //ExFor:SwfSaveOptions
            //ExFor:SwfSaveOptions.ToolTipsFontName
            //ExFor:SwfSaveOptions.ToolTips
            //ExFor:SwfViewerControlIdentifier
            //ExSummary:Shows how to change the the tooltips used in the embedded document viewer.
            // We create an instance of SwfSaveOptions to specify our custom tooltips.
            SwfSaveOptions options = new SwfSaveOptions();

            // By default, all tooltips are in English. You can specify font used for each tooltip.
            // Note that font specified should contain proper glyphs normally used in tooltips.
            options.ToolTipsFontName = "Times New Roman";

            // The following code will set the tooltip used for each control. In our case we will change the tooltips from English
            // to Russian.
            options.ToolTips[SwfViewerControlIdentifier.TopPaneActualSizeButton] = "Оригинальный размер";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneFitToHeightButton] = "По высоте страницы";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneFitToWidthButton] = "По ширине страницы";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneZoomOutButton] = "Увеличить";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneZoomInButton] = "Уменшить";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneSelectionModeButton] = "Режим выделения текста";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneDragModeButton] = "Режим перетаскивания";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneSinglePageScrollLayoutButton] = "Одностнаничный скролинг";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneSinglePageLayoutButton] = "Одностраничный режим";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneTwoPageScrollLayoutButton] = "Двустраничный скролинг";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneTwoPageLayoutButton] = "Двустраничный режим";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneFullScreenModeButton] = "Полноэкранный режим";
            options.ToolTips[SwfViewerControlIdentifier.TopPanePreviousPageButton] = "Предыдущая старница";
            options.ToolTips[SwfViewerControlIdentifier.TopPanePageField] = "Введите номер страницы";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneNextPageButton] = "Следующая страница";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneSearchField] = "Введите искомый текст";
            options.ToolTips[SwfViewerControlIdentifier.TopPaneSearchButton] = "Искать";
            
            // Left panel.
            options.ToolTips[SwfViewerControlIdentifier.LeftPaneDocumentMapButton] = "Карта документа";
            options.ToolTips[SwfViewerControlIdentifier.LeftPanePagePreviewPaneButton] = "Предварительный просмотр страниц";
            options.ToolTips[SwfViewerControlIdentifier.LeftPaneAboutButton] = "О приложении";
            options.ToolTips[SwfViewerControlIdentifier.LeftPaneCollapsePanelButton] = "Свернуть панель";
            
            // Bottom panel.
            options.ToolTips[SwfViewerControlIdentifier.BottomPaneShowHideBottomPaneButton] = "Показать/Скрыть панель";
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\SwfSaveOptions.ToolTips.swf", options);
        }

        [Test]
        public void HideControls()
        {
            //ExStart
            //ExFor:SwfSaveOptions.TopPaneControlFlags
            //ExFor:SwfTopPaneControlFlags
            //ExFor:SwfSaveOptions.ShowSearch
            //ExSummary:Shows how to choose which controls to display in the embedded document viewer.
            Document doc = new Document(MyDir + "Document.doc");

            // Create an instance of SwfSaveOptions and set some buttons as hidden.
            SwfSaveOptions options = new SwfSaveOptions();
            // Hide all the buttons with the exception of the page control buttons. Similar flags can be used for the left control pane as well.
            options.TopPaneControlFlags = SwfTopPaneControlFlags.HideAll | SwfTopPaneControlFlags.ShowActualSize |
                SwfTopPaneControlFlags.ShowFitToWidth | SwfTopPaneControlFlags.ShowFitToHeight |
                SwfTopPaneControlFlags.ShowZoomIn | SwfTopPaneControlFlags.ShowZoomOut;

            // You can also choose to show or hide the main elements of the viewer. Hide the search control.
            options.ShowSearch = false;
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\SwfSaveOptions.HideControls.swf", options);
        }

        [Test]
        public void SetLogo()
        {
            Document doc = new Document(MyDir + "Document.doc");

            //ExStart
            //ExFor:SwfSaveOptions.#ctor
            //ExFor:SwfSaveOptions
            //ExFor:SwfSaveOptions.LogoImageBytes
            //ExFor:SwfSaveOptions.LogoLink
            //ExSummary:Shows how to specify a custom logo and link it to a web address in the embedded document viewer.
            // Create an instance of SwfSaveOptions.
            SwfSaveOptions options = new SwfSaveOptions();

            // Read the image into byte array.
            byte[] logoBytes = File.ReadAllBytes(MyDir + "LogoSmall.png");

            // Specify the logo image to use.
            options.LogoImageBytes = logoBytes;

            // You can specify the URL of web page that should be opened when you click on the logo.
            options.LogoLink = "http://www.aspose.com";
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\SwfSaveOptions.CustomLogo.swf", options);
        }

    }
}
