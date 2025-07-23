// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExLicense : ApiExampleBase
    {
        [Test]
        public void LicenseFromFileNoPath()
        {
            //ExStart
            //ExFor:License
            //ExFor:License.#ctor
            //ExFor:License.SetLicense(String)
            //ExSummary:Shows how to initialize a license for Aspose.Words using a license file in the local file system.
            
#if CPLUSPLUS
            string testLicenseFileName = "Aspose.Total.C++.lic";
#else
            string testLicenseFileName = "Aspose.Total.NET.lic";
#endif
            
            // Set the license for our Aspose.Words product by passing the local file system filename of a valid license file.
            string licenseFileName = Path.Combine(LicenseDir, testLicenseFileName);

            License license = new License();
            license.SetLicense(licenseFileName);

            // Create a copy of our license file in the binaries folder of our application.
            string licenseCopyFileName = Path.Combine(AssemblyDir, testLicenseFileName);
            File.Copy(licenseFileName, licenseCopyFileName);

            // If we pass a file's name without a path,
            // the SetLicense will search several local file system locations for this file.
            // One of those locations will be the "bin" folder, which contains a copy of our license file.
            license.SetLicense(testLicenseFileName);
            //ExEnd

            license.SetLicense("");
            File.Delete(licenseCopyFileName);
        }

        [Test]
        public void LicenseFromStream()
        {
            //ExStart
            //ExFor:License.SetLicense(Stream)
            //ExSummary:Shows how to initialize a license for Aspose.Words from a stream.
#if CPLUSPLUS
            string testLicenseFileName = "Aspose.Total.C++.lic";
#else
            string testLicenseFileName = "Aspose.Total.NET.lic";
#endif
            // Set the license for our Aspose.Words product by passing a stream for a valid license file in our local file system.
            using (Stream myStream = File.OpenRead(Path.Combine(LicenseDir, testLicenseFileName)))
            {
                License license = new License();
                license.SetLicense(myStream);
            }
            //ExEnd
        }
    }
}
