// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    /// <summary>
    /// Provides common infrastructure for all API examples that are implemented as unit tests.
    /// </summary>
    public class ApiExampleBase
    {
        private static readonly string mExternalAppPath = Android.OS.Environment.ExternalStorageDirectory.AbsolutePath;
        
        [SetUp]
        public void SetUp()
        {
            SetUnlimitedLicense();

            if (Directory.Exists(ArtifactsDir))
            {
                try
                {
                    Directory.Delete(ArtifactsDir, true);
                    Directory.CreateDirectory(ArtifactsDir);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }
                
            }
            else
            {
                try
                {
                    Directory.CreateDirectory(ArtifactsDir);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }
            }
        }

        [TearDown]
        public void TearDown()
        {
            try
            {
                Directory.Delete(ArtifactsDir, true);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        internal static void SetUnlimitedLicense()
        {
            if (File.Exists(TestLicenseFileName))
            {
                // This shows how to use an Aspose.Words license when you have purchased one.
                // You don't have to specify full path as shown here. You can specify just the 
                // file name if you copy the license file into the same folder as your application
                // binaries or you add the license to your project as an embedded resource.
                License license = new License();
                license.SetLicense(TestLicenseFileName);
            }
        }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string ArtifactsDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string MyDir { get; }

        /// <summary>
        /// Gets the path to the images used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string ImageDir { get; }

        /// <summary>
        /// Gets the path of the demo database. Ends with a back slash.
        /// </summary>
        internal static string DatabaseDir { get; }

        /// <summary>
        /// Gets the path of the free fonts. Ends with a back slash.
        /// </summary>
        internal static string FontsDir { get; }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string GoldsDir { get; }

        /// <summary>
        /// Gets the url of the Aspose logo.
        /// </summary>
        internal static string AsposeLogoUrl { get; }

        static ApiExampleBase()
        {
            ArtifactsDir = Path.Combine(mExternalAppPath, "Data/Artifacts/");
            MyDir = Path.Combine(mExternalAppPath, "Data/");
            ImageDir = Path.Combine(mExternalAppPath, "Data/Images/");
            DatabaseDir = Path.Combine(mExternalAppPath, "Data/Database/");
            GoldsDir = Path.Combine(mExternalAppPath, "Data/Golds/");
            FontsDir = Path.Combine(mExternalAppPath, "Data/MyFonts/");
            AsposeLogoUrl = new Uri("https://www.aspose.cloud/templates/aspose/App_Themes/V3/images/words/header/aspose_words-for-net.png").AbsoluteUri;
        }

        internal static readonly string TestLicenseFileName = Path.Combine(mExternalAppPath, "Data/License/Aspose.Words.lic");
    }
}