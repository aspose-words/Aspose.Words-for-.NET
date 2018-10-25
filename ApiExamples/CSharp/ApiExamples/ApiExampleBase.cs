﻿// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    /// <summary>
    /// Provides common infrastructure for all API examples that are implemented as unit tests.
    /// </summary>
    public class ApiExampleBase
    {
        private readonly string artifactsPath = MyDir + @"\Artifacts\";

        [SetUp]
        public void SetUp()
        {
            if (!CheckForSkipSetUp())
            {
                SetUnlimitedLicense();

                if (!Directory.Exists(artifactsPath))
                    //Create new empty directory
                    Directory.CreateDirectory(artifactsPath);
            }
        }

        [TearDown]
        public void TearDown()
        {
            if (!CheckForSkipTearDown())
            {
                //Delete all dirs and files from directory
                Directory.Delete(artifactsPath, true);
            }
        }

        /// <summary>
        /// Checks when we need to skip precondition before test.
        /// </summary>
        private static bool CheckForSkipSetUp()
        {
            bool skipSetup = TestContext.CurrentContext.Test.Properties["Category"].Contains("SkipSetup");
            return skipSetup;
        }

        /// <summary>
        /// Checks when we need to skip post-condition after test.
        /// </summary>
        private static bool CheckForSkipTearDown()
        {
            bool skipSetup = TestContext.CurrentContext.Test.Properties["Category"].Contains("SkipTearDown");
            return skipSetup;
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
        /// Returns the assembly directory correctly even if the assembly is shadow-copied.
        /// </summary>
        internal static string GetAssemblyDir(Assembly assembly)
        {
            // CodeBase is a full URI, such as file:///x:\blahblah.
            Uri uri = new Uri(assembly.CodeBase);
            return Path.GetDirectoryName(uri.LocalPath) + Path.DirectorySeparatorChar;
        }

        /// <summary>
        /// Gets the path to the currently running executable.
        /// </summary>
        internal static string AssemblyDir
        {
            get { return gAssemblyDir; }
        }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string MyDir
        {
            get { return gMyDir; }
        }

        /// <summary>
        /// Gets the path to the images used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string ImageDir
        {
            get { return gImageDir; }
        }

        /// <summary>
        /// Gets the path of the demo database. Ends with a back slash.
        /// </summary>
        internal static string DatabaseDir
        {
            get { return gDatabaseDir; }
        }

        static ApiExampleBase()
        {
            gAssemblyDir = GetAssemblyDir(Assembly.GetExecutingAssembly());
            gMyDir = new Uri(new Uri(gAssemblyDir), @"../../../../Data/").LocalPath;
            gImageDir = new Uri(new Uri(gAssemblyDir), @"../../../../Data/Images/").LocalPath;
            gDatabaseDir = new Uri(new Uri(gAssemblyDir), @"../../../../Data/Database/").LocalPath;
        }

        internal static readonly string gAssemblyDir;
        internal static readonly string gMyDir;
        internal static readonly string gImageDir;
        internal static readonly string gDatabaseDir;

        /// <summary>
        /// This is where the test license is on my development machine.
        /// </summary>
        internal const string TestLicenseFileName = @"X:\awnet\TestData\Licenses\Aspose.Total.lic";
    }
}