using System;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
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
            Console.WriteLine("I am currently using {0}, version number {1}.", BuildVersionInfo.Product, BuildVersionInfo.Version);
            //ExEnd
        }
    }
}
