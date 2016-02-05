using System;
using NUnit.Framework;


namespace ApiExamples.BuildVersionInfo
{
    [TestFixture]
    public class ExBuildVersion : ApiExampleBase
    {
        [Test]
        public void ShowBuildVersionInfo()
        {
            //ExStart
            //ExFor:BuildVersionInfo
            //ExSummary:Shows how to use BuildVersionInfo to obtain information about this product.
            Console.WriteLine("I am currently using {0}, version number {1}.", Aspose.Words.BuildVersionInfo.Product, Aspose.Words.BuildVersionInfo.Version);
            //ExEnd
        }
    }
}
