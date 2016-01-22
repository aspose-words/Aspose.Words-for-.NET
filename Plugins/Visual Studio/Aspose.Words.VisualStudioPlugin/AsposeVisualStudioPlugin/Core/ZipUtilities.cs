// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Ionic.Zip;
namespace AsposeVisualStudioPluginWords.Core
{
    public class ZipUtilities
    {


        /// <summary>
        /// Unzip files
        /// </summary>
        /// <param name="zipFilePath"></param>
        /// <param name="pathToExtract"></param>
        public Boolean ExtractZipFile(string zipFilePath, string pathToExtract)
        {
            try
            {
                var options = new ReadOptions { StatusMessageWriter = System.Console.Out };
                using (ZipFile zip = ZipFile.Read(zipFilePath, options))
                {
                    // This call to ExtractAll() assumes:
                    //   - none of the entries are password-protected.
                    //   - want to extract all entries to current working directory
                    //   - none of the files in the zip already exist in the directory;
                    //     if they do, the method will throw.
                    zip.ExtractAll(pathToExtract);
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
  
    }
}
