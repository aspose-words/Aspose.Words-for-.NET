==========================================
Aspose.Words for .NET Samples Read Me
==========================================

This package contains C# and VB.NET Sample Projects for Aspose.Words for .NET.

The Sample Projects are additional projects that are distributed separately from the Aspose.Words for .NET itself.

The solution files of Microsoft Visual Studio 2005, 2008 and 2010 are provided. 

Please open the solution file (.sln) according to the version of Visual Studio installed on your system.

Aspose.Words.Samples.2005.sln - Open with Microsoft Visual Studio 2005
Aspose.Words.Samples.2008.sln - Open with Microsoft Visual Studio 2008
Aspose.Words.Samples.2010.sln - Open with Microsoft Visual Studio 2010

How to Install
==========================================

If the installer has been used to install Aspose.Words for .NET then the samples in this archive can be extracted to any location and they should run correctly without any other actions. 

Otherwise the "Aspose.Words for .NET Dlls Only" archive should be downloaded and extracted to any location and the sample folders from this archive extracted along side in a folder called "Samples". The Aspose.Cells and Aspose.Network libraries are also required for a few of the samples.


How to Run the Demos
==========================================

Open the appropriate solution file (.sln) in Microsoft Visual Studio. Click on "Debug" menu and choose one of the following menu items:

- Start Debugging
- Start Without Debugging


Software Requirements
==========================================

- Aspose.Words for .NET 10.6.0 or later (Remove any Aspose.Words for .NET assembly from GAC to get rid of possible errors)
- Additional libraries Aspose.Network for .NET 6.4 and Apose.Cells for .NET 6.0 or later are required to compile some sample projects that  demonstrate integration between several Aspose products. These libraries are shipped with the MSI installer, if this has been used then no extra actions are required. In any other case these libraries must be downloaded manually from the Aspose site.
- Some samples require extra frameworks:
  - There are samples which demonstrate integration with the Silverlight or Ajax frameworks.
  - The Examples project requires NUnit framework to be installed. Each code example within this project is structured as a test which can be executed automatically using the NUnit framework. NUnit can be downloaded from http://www.nunit.org/
- Any of the following Internet Browsers  
  - Microsoft Internet Explorer 6 or above
  - Mozilla Firefox 3 or above
  - Google Chrome
- Any of the following versions of Microsoft Visual Studio
  - Microsoft Visual Studio 2005 SP1 (Express, Standard, Professional and Team Editions)
     - Note that some samples which require .NET 3.5 features may not compile under Microsoft Visual Studio 2005.
  - Microsoft Visual Studio 2008 SP1 (Express, Standard, Professional and Team Editions)
  - Microsoft Visual Studio 2010 (Express, Professional, Premium and Ultimate Editions)


IIS Requirements
==========================================

All demos can run with the ASP.NET Development Server installed with Visual Studio.


Running the Samples under Linux and MacOS (using Mono)
==========================================

The Aspose.Words for .NET samples have been tested on Linux (Ubuntu 11.04) and using the Mono Framework (2.10). There is a good chance that the samples will work well under other operating systems as well but these have yet to be tested.

Some samples may contain features that are not yet available on the Mono Framework, for instance SilverLight, Ajax and mailing technologies. These therefore will not compile or run as expected.

For the samples to work properly you should copy the "Samples" folder to the target machine so it's placed at that same level as the Aspose.Words "Bin" folder from the Aspose.Words.zip archive. This folder contain the Aspose.Words library and placing the package here allows the samples to correctly reference the Aspose.Words library. In addition, in order for all samples to work you are also required to acquire the libraries for "Aspose.Cells" and "Aspose.Network" from the Aspose site.

Alternatively, you can run the samples from any location on the machine by running the "Linux-Register-Dll-With-MonoDevelop.sh" script found in the Dlls Only archive. Note that this script must be run with root privileges. This will add Aspose.Words to the GAC and add an assembly reference to the appropriate location within MonoDevelop. Any extra Dlls required can be downloaded and located manually. 


Software Requirements for Running Demos on Linux
==========================================

- The Mono Framework installed
- The "libgda" drivers package installed.

Currently there are issues with OleConnection on Linux which is used to access databases in some demos. This causes certain demos which use the connection to fail with an exception stating that a required driver is missing. More details on what causes this are being investigated and this issue updated as soon as possible. 


http://www.aspose.com/
Copyright (c) 2001-2011 Aspose Pty Ltd. All Rights Reserved.