# Word Document Processing .NET API

![Version 21.9](https://img.shields.io/badge/nuget-v21.9-blue) ![Nuget](https://img.shields.io/nuget/dt/Aspose.Words)

[Product Page](https://products.aspose.com/words/net/) | [Docs](https://docs.aspose.com/words/net/) | [Demos](https://products.aspose.app/words/family) | [API Reference](https://apireference.aspose.com/words/net) | [Examples](https://github.com/aspose-words/Aspose.Words-for-.NET/tree/master/Examples) | [Blog](https://blog.aspose.com/category/words/) | [Search](https://search.aspose.com/) | [Free Support](https://forum.aspose.com/c/words) | [Temporary License](https://purchase.aspose.com/temporary-license)

[Aspose.Words for .NET](https://products.aspose.com/words/net/) is a class library that can be used by C#, F#, VB.NET developers for a variety of document-processing tasks, including document generation, modification, converting, and rendering. Our library is self-sufficient and doesn't depend on any third-party software, such as Microsoft Word, OpenOffice, and similar office suites. 

This package can be used to develop applications for a vast range of operating systems (Windows, Linux, macOS, iOS, Android) and [platforms](https://docs.aspose.com/words/net/supported-platforms/) such as [Windows Azure](https://docs.aspose.com/words/net/windows-azure-platform/), Xamarin.Android, Xamarin.iOS, Xamarin.Mac. You can build both 32-bit and 64-bit software, including ASP.NET, WCF, and WinForms. Also, you can use our library via COM Interop from ASP, PHP, Perl, and Python programming languages.

*Please note*: our library implies the use of [.NET programming languages](https://en.wikipedia.org/wiki/List_of_CLI_languages), compatible with CLI/CLS technical standards. If you require a corresponding native library for C++, you can download it from [here](https://www.nuget.org/packages/Aspose.Words.Cpp/).

## Functionality

- Provides comprehensive [document import and export](https://docs.aspose.com/words/net/loading-saving-and-converting/) with [35+ supported file formats](https://docs.aspose.com/words/net/supported-document-formats/). This allows developers to [convert documents](https://docs.aspose.com/words/net/convert-a-document/) from [one file format](https://apireference.aspose.com/words/net/aspose.words/loadformat) to [another](https://apireference.aspose.com/words/net/aspose.words/saveformat). For example, you can convert PDF to Word and Word to PDF documents with professional quality.
- Provides full access to all Word and OpenOffice document elements, including formatting properties and styling.
- Provides [high-fidelity rendering](https://docs.aspose.com/words/net/rendering/) of Word documents to PDF, JPG, PNG and other imaging formats.
- Provides the ability to [print](https://docs.aspose.com/words/net/print-a-document-programmatically-or-using-dialogs/) OpenOffice and Word documents programmatically."
- Supports powerful [Report Generation with Mail Merge](https://docs.aspose.com/words/net/mail-merge-and-reporting/) functionality, which allows to create documents dynamically using templates and data sources.
- Contains a flexible [LINQ Reporting Engine](https://docs.aspose.com/words/net/linq-reporting-engine/), designed to fetch data from databases, XML, JSON, OData, external documents.
- Provides a rich set of utility functions, you can use to [split a document](https://docs.aspose.com/words/net/split-a-document/) into parts, [join documents together](https://docs.aspose.com/words/net/insert-and-append-documents/), [compare documents](https://docs.aspose.com/words/net/compare-documents/), and much more.

To become familiar with the most popular Aspose.Words functionality, please have a look at our [free online applications](https://products.aspose.app/words/family).


## Supported Formats
### Read and Write Formats

**Microsoft Word:** DOC, DOT, DOCX, DOTX, DOTM, FlatOpc, FlatOpcMacroEnabled, FlatOpcTemplate, FlatOpcTemplateMacroEnabled, RTF, WordML\
**OpenDocument:** ODT, OTT\
**Web:** HTML, MHTML\
**Markdown:** MD\
**Fixed Layout:** PDF\
**Text:** TXT

### Read-Only Formats

**Microsoft Word:** DocPreWord60\
**eBook:** MOBI, CHM

### Write-Only Formats

**Fixed Layout:** XPS, OpenXps\
**PostScript:** PS\
**Printer:** PCL\
**eBook:** EPUB\
**Markup:** XamlFixed, HtmlFixed, XamlFlow, XamlFlowPack\
**Image:** SVG, TIFF, PNG, BMP, JPEG, GIF\
**Metafile:** EMF

## Getting Started

So, you probably want to jump up and start coding your document processing application on C#, F# or Visual Basic right away? Let us show you how to do it in a few easy steps.

Run ```Install-Package Aspose.Words``` from the Package Manager Console in Visual Studio to fetch the NuGet package.
If you want to upgrade to the latest package version, please run ```Update-Package Aspose.Words```.

You can run the following code snippets in C# to see how our library works. Also feel free to check out the [GitHub Repository](https://github.com/aspose-words/Aspose.Words-for-.NET) for other common use cases.

### Create a DOCX using C#

Aspose.Words for .NET allows you to create a blank Word document and add content to the file.

```c#
// Create a Word document.
Document doc = new Document();

// Use a DocumentBuilder instance to add content to the file.
DocumentBuilder builder = new DocumentBuilder(doc);

// Write a new paragraph to the document.
builder.Writeln("This is an example of a Word document created in C#");

// Save it as a DOCX file. The output format is automatically determined by the filename extension.
doc.Save(dir + "OutputWordDocument.docx");
```

### Create a PDF in C#

Aspose.Words for .NET allows you to create a new PDF file and fill it with data.

```c#
// Create a PDF document.
Document pdf = new Document();

// Use a DocumentBuilder instance to add content to the file.
DocumentBuilder builder = new DocumentBuilder(pdf);

// Write a new paragraph to the document.
builder.Writeln("This is an example of a PDF document created using C#");

// Save it as a PDF file.
pdf.Save(dir + "OutputDocument.pdf");
```

### Convert a Word document to HTML with C#

You can convert Microsoft Word to PDF, XPS, Markdown, HTML, JPEG, TIFF, and other file formats. The following snippet demonstrates the conversion from DOCX to HTML:

```c#
// Load a Word file from the local drive.
Document doc = new Document(dir + "InputWordDocument.docx");

// Save it to HTML format.
doc.Save(dir + "OutputHtmlDocument.html");
```

### Import a PDF and save as a DOCX via C#

In addition, you can import a PDF file into your .NET application and export it as a DOCX format file without the need to install Microsoft Word:

```c#
// Load a PDF file from the local drive.
Document pdf = new Document(dir + "InputDocument.pdf");

// Save it to DOCX format.
pdf.Save(dir + "OutputWordDocument.docx");
```

[Product Page](https://products.aspose.com/words/net/) | [Docs](https://docs.aspose.com/words/net/) | [Demos](https://products.aspose.app/words/family) | [API Reference](https://apireference.aspose.com/words/net) | [Examples](https://github.com/aspose-words/Aspose.Words-for-.NET/tree/master/Examples) | [Blog](https://blog.aspose.com/category/words/) | [Search](https://search.aspose.com/) | [Free Support](https://forum.aspose.com/c/words) | [Temporary License](https://purchase.aspose.com/temporary-license)
