using System;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.BuildVersionInfo
{
    [TestFixture]
    public class ExBuildVersion : QaTestsBase
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
