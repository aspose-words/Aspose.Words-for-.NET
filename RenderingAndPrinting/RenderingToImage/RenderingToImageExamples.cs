using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.RenderingAndPrinting.RenderingToImage
{
    [TestClass, TestFixture]
    public class RenderingToImageExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void SaveAsMultipageTiff()
        {
            TestHelper.SetDataDir("RenderingAndPrinting/RenderingToImage/SaveAsMultipageTiff");

            SaveAsMultipageTiffExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void RenderShapes()
        {
            TestHelper.SetDataDir("RenderingAndPrinting/RenderingToImage/RenderShapes");

            RenderShapesExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void ImageColorFilters()
        {
            TestHelper.SetDataDir("RenderingAndPrinting/RenderingToImage/ImageColorFilters");

            ImageColorFiltersExample.Program.Main();
        }

    }
}
