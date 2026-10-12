# Summary: Big document.docx

An introductory product guide for **Aspose.Words for .NET**, a class library for generating, modifying, converting, rendering and printing documents without Microsoft Word.

## Getting Started
- **Supported formats:** DOC, DOCX, RTF, HTML, OpenDocument, PDF, XPS, EPUB and others.
- **Four feature areas:** conversions, a Document Object Model for creating and editing content, rendering (to PDF, XPS, images and .NET Graphics), and reporting (filling templates from data sources).
- **No Word automation:** Microsoft Office does not need to be installed. The guide presents this as more secure, stable, scalable and cheaper than automating Word.
- **Platform independence:** the .NET and Java editions are "twin" products that run on Windows, Linux and Mac OS. The .NET edition also works via COM Interop and on Mono.
- **Performance:** it ships as a single assembly. It is thread-safe as long as each document is handled by only one thread at a time.
- **Learning curve:** there are 150+ public classes, but the names follow the Word Automation model and the document is a tree of objects, much like an XML DOM.

## Features
- **High-quality conversions:** converting a document takes two lines of code (load a `Document`, call `Save`). C# and VB examples are included.
- **Fidelity:** DOC, DOCX, RTF and WordprocessingML convert with "high fidelity", meaning content and formatting are fully kept. HTML conversion is very good but not high-fidelity, because some features (such as headers/footers, sections and fields) have no HTML equivalent.
- **Scalability and speed:** a DOC file over 5 MB loads in under a second, and small files can be processed thousands per minute.
- **DOC format:** reads and writes files from Word 97–2010, Word for Mac and OpenOffice. It fully supports advanced features such as revisions, shapes, OLE/ActiveX objects, VBA, embedded fonts, encryption and right-to-left text.
- **WordprocessingML/DOCX:** supports Word 2003 XML, OOXML ECMA-376 and ISO/IEC 29500, and Flat OPC.

## Why use a library for open OOXML?
The OOXML specification runs to thousands of pages. Even a simple find-and-replace is hard to do by hand, because text is split across formatting runs (shown with a "Hello World" XML fragment). Aspose.Words provides conversions, merging content from several documents, one-line operations (field update, accepting revisions, mail merge) and find/replace across ranges. The document's conclusion: "open and standard does not mean simple."
