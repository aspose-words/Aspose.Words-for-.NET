// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExBuildVersion : ApiExampleBase
    {
        [Test]
        public void PrintBuildVersionInfo()
        {
            //ExStart
            //ExFor:BuildVersionInfo
            //ExFor:BuildVersionInfo.Product
            //ExFor:BuildVersionInfo.Version
            //ExSummary:Shows how to display information about your installed version of Aspose.Words.
            Console.WriteLine($"I am currently using {BuildVersionInfo.Product}, version number {BuildVersionInfo.Version}!");
            //ExEnd

            Assert.AreEqual("Aspose.Words for .NET", BuildVersionInfo.Product);
            Assert.True(Regex.IsMatch(BuildVersionInfo.Version, "[0-9]{2}.[0-9]{1,2}"));
        }
    }
}