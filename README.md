![Nuget](https://img.shields.io/nuget/v/Aspose.Words) ![Nuget](https://img.shields.io/nuget/dt/Aspose.Words) ![GitHub](https://img.shields.io/github/license/aspose-words/Aspose.Words-for-.NET)
# Word Processing API for .NET

[Aspose.Words for .NET](https://products.aspose.com/words/net) is a powerful on-premise class library that can be used for numerous document processing tasks. It enables developers to enhance their own applications with features such as generating, modifying, converting, rendering, and printing documents, without relying on third-party applications, for example, Microsoft Word, or Office Automation.

<p align="center">

  <a title="Download complete Aspose.Words for .NET source code" href="https://github.com/aspose-words/Aspose.Words-for-.NET/archive/master.zip">
	<img src="https://raw.github.com/AsposeExamples/java-examples-dashboard/master/images/downloadZip-Button-Large.png" />
  </a>
</p>

This repository contains [Demos](Demos), [Examples](Examples), [Plugins](Plugins) and [Showcases](Showcases) for [Aspose.Words for .NET](http://www.aspose.com/products/words/net) to help you learn and write your own applications.

Directory | Description
--------- | -----------
[Demos](Demos)  | Aspose.Words for .NET Live Demos Source Code
[Examples](Examples)  | A collection of .NET examples that help you learn and explore the API features
[Showcases](Showcases)  | Standalone ready-to-use applications that demonstrate some specific use cases
[Plugins](Plugins)  | Plugins that will demonstrate one or more features of Aspose.Words for .NET

# .NET API for Various Document Formats

[Aspose.Words for .NET](https://products.aspose.com/words/net) is a powerful on-premise class library that can be used for numerous document processing tasks. It enables developers to enhance their own applications with features such as generating, modifying, converting, rendering, and printing documents, without relying on third-party applications, for example, Microsoft Word, or automation.

## Word API Features

- Aspose.Words can be used to develop applications for a vast range of operating systems such as Windows or Linux & Mac OS and [platforms](https://docs.aspose.com/display/wordsnet/Feature+Overview#FeatureOverview-SupportedPlatforms) such as [Windows Azure](https://docs.aspose.com/display/wordsnet/Windows+Azure+Platform), Xamarin.Android, or Xamarin.iOS&Xamarin.Mac.
- Comprehensive [document import and export](https://docs.aspose.com/display/wordsnet/Loading%2C+Saving+and+Converting) with [35+ supported file formats](https://docs.aspose.com/display/wordsnet/Supported+Document+Formats). This allows users to convert documents from one popular format to another, for example, from DOCX into PDF or Markdown, or from PDF into various Word formats.
- Programmatic access to the formatting properties of all document elements. For example, using Aspose.Words users can [split a document](https://docs.aspose.com/display/wordsnet/Split+a+Document) into parts or [compare two documents](https://docs.aspose.com/display/wordsnet/Compare+Documents).
- [High fidelity rendering](https://docs.aspose.com/display/wordsnet/Rendering) of document pages. For example, if it is needed to render a document as in Microsoft Word, Aspose.Words will successfully cope with this task.
- Ability to [print a document programmatically](https://docs.aspose.com/display/wordsnet/Print+a+Document+Programmatically+or+Using+Dialogs) using Aspose.Words and the XpsPrint API or via dialog boxes.
- [Generate reports with Mail Merge](https://docs.aspose.com/display/wordsnet/Mail+Merge+and+Reporting), which allows filling in merge templates with data from various sources to create merged documents.
- [LINQ Reporting Engine](https://docs.aspose.com/display/wordsnet/LINQ+Reporting+Engine) to fetch data from databases, XML, JSON, OData, external documents, and much more.

## Read & Write Document Formats

**Microsoft Word:** DOC, DOCX, RTF, DOT, DOTX, DOTM, DOCM FlatOPC, FlatOpcMacroEnabled, FlatOpcTemplate, FlatOpcTemplateMacroEnabled\
**OpenOffice:** ODT, OTT\
**WordprocessingML:** WordML\
**Web:** HTML, MHTML\
**Fixed Layout:** PDF\
**Text:** TXT

## Save Word Files As

**Fixed Layout:** XPS, OpenXPS, PostScript (PS)\
**Images:** TIFF, JPEG, PNG, BMP, SVG, EMF, GIF\
**Web:** HtmlFixed\
**Others:** PCL, EPUB, XamlFixed, XamlFlow, XamlFlowPack

## Platform Independence

Aspose.Words for .NET API can be used to develop applications for a vast range of operating systems (Windows, Linux & Mac OS) and platforms. You can build both 32-bit and 64-bit applications including ASP.NET, WCF & WinForms. Aspose.Words for .NET can also be used via COM Interop from ASP, PHP, Perl and Python. 

You can also build applications with Mono as well as on Windows Azure, Microsoft SharePoint, Microsoft Silverlight, Xamarin.Android, Xamarin.iOS & Xamarin.Mac.

## Getting Started with Aspose.Words for .NET

Ready to give Aspose.Words for .NET a try? Simply run `Install-Package Aspose.Words` from Package Manager Console in Visual Studio to fetch the NuGet package. If you already have Aspose.Words for .NET and want to upgrade the version, please run `Update-Package Aspose.Words` to get the latest version.

## Using C# to Create a DOC File from Scratch

You can execute this snippet in your environment to see how Aspose.Words performs or check the [GitHub Repository](https://github.com/aspose-words/Aspose.Words-for-.NET) for other common usage scenarios.

```csharp
// create a blank document
Document doc = new Document();
// the DocumentBuilder class provides members to easily add content to a document
DocumentBuilder builder = new DocumentBuilder(doc);
// write a new paragraph in the document with the text "Hello World!"
builder.Writeln("Hello World!");
// save the document in DOCX format. 
// the format to save as is inferred from the extension of the file name.
doc.Save(dir + "output.docx");
```

## Using C# to Export DOC to EPUB Format

Aspose.Words for .NET allows you to convert Microsoft Word® formats into bytes, HTML, EPUB, MHTML and other file formats. Following snippet demonstrates the conversion of a DOC file to an EPUB file:

```csharp
// load the document from disc
Document doc = new Document(dir + "template.doc");
// save the document in EPUB format
doc.Save(dir + "output.epub");
```

## Using C# to Import PDF and Save as a DOCX File

The following code sample demonstrates, how you can import a PDF document into your .NET application and export it as a DOCX format file without the need to install the Microsoft Word®:

```csharp
// Load the PDF document from directory
Document doc = new Document(dir + "input.pdf");

// Save the document in DOCX format
doc.Save(dir + "output.docx");
```

[Product Page](https://products.aspose.com/words/net) | [Docs](https://docs.aspose.com/display/wordsnet/Home) | [Demos](https://products.aspose.app/words/family) | [API Reference](https://apireference.aspose.com/words/net) | [Examples](https://github.com/aspose-words/Aspose.Words-for-.NET) | [Blog](https://blog.aspose.com/category/words/) | [Free Support](https://forum.aspose.com/c/words) | [Temporary License](https://purchase.aspose.com/temporary-license)

