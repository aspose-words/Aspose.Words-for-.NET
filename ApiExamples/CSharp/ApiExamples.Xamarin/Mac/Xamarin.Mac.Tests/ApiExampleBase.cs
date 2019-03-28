// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
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
        private static readonly string mExternalAppPath = "/Users/vyacheslav/falleretic.Aspose.Words-for-.NET/ApiExamples/";

        [SetUp]
        public void OneTimeSetUp()
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
        public void OneTimeTearDown()
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
        internal static String ArtifactsDir
        {
            get { return gArtifactsDir; }
        }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static String MyDir
        {
            get { return gMyDir; }
        }

        /// <summary>
        /// Gets the path to the images used by the code examples. Ends with a back slash.
        /// </summary>
        internal static String ImageDir
        {
            get { return gImageDir; }
        }

        /// <summary>
        /// Gets the path of the demo database. Ends with a back slash.
        /// </summary>
        internal static String DatabaseDir
        {
            get { return gDatabaseDir; }
        }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static String GoldsDir
        {
            get { return gGoldsDir; }
        }
		
		/// <summary>
        /// Gets the url of the Aspose logo.
        /// </summary>
        internal static string AsposeLogoUrl { get; }

        static ApiExampleBase()
        {
            gArtifactsDir = Path.Combine(mExternalAppPath, "Data/Artifacts/");
            gMyDir = Path.Combine(mExternalAppPath, "Data/");
            gImageDir = Path.Combine(mExternalAppPath, "Data/Images/");
            gDatabaseDir = Path.Combine(mExternalAppPath, "Data/Database/");
            gGoldsDir = Path.Combine(mExternalAppPath, "Data/Golds/");
			AsposeLogoUrl = new Uri("https://www.aspose.cloud/templates/aspose/App_Themes/V3/images/words/header/aspose_words-for-net.png").AbsoluteUri;
        }

        private static readonly String gArtifactsDir;
        private static readonly String gMyDir;
        private static readonly String gImageDir;
        private static readonly String gDatabaseDir;
        private static readonly String gGoldsDir;

        internal static readonly string TestLicenseFileName = Path.Combine(mExternalAppPath, "Data/License/Aspose.Words.lic");
    }
}