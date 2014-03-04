Imports System.Reflection
Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

<Assembly: AssemblyTitle("Aspose.Words Examples")>
<Assembly: AssemblyDescription("A collection of examples which demonstrate how to use the Aspose.Words for .NET API.")>
<Assembly: AssemblyConfiguration("VisualBasic")>

Namespace Examples.QuickStart
	<TestClass, TestFixture> _
	Public Class QuickStartExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub HelloWorld()
	            TestHelper.SetDataDir("QuickStart/HelloWorld")
	
	            HelloWorldExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub AppendDocuments()
	            TestHelper.SetDataDir("QuickStart/AppendDocuments")
	
	            AppendDocumentsExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub ApplyLicense()
	            TestHelper.SetDataDir("QuickStart/ApplyLicense")
	
	            ApplyLicenseExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub FindAndReplace()
	            TestHelper.SetDataDir("QuickStart/FindAndReplace")
	
	            FindAndReplaceExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub SimpleMailMerge()
	            TestHelper.SetDataDir("QuickStart/SimpleMailMerge")
	
	            SimpleMailMergeExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub LoadAndSaveToDisk()
	            TestHelper.SetDataDir("QuickStart/LoadAndSaveToDisk")
	
	            LoadAndSaveToDiskExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub LoadAndSaveToStream()
	            TestHelper.SetDataDir("QuickStart/LoadAndSaveToStream")
	
	            LoadAndSaveToStreamExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub Doc2Pdf()
	            TestHelper.SetDataDir("QuickStart/Doc2Pdf")
	
	            Doc2PdfExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub UpdateFields()
	            TestHelper.SetDataDir("QuickStart/UpdateFields")
	
	            UpdateFieldsExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub WorkingWithNodes()
	            TestHelper.SetDataDir("QuickStart/WorkingWithNodes")
	
	            WorkingWithNodesExample.Program.Main()
        End Sub

	End Class

	<TestClass, SetUpFixture> _
	Public Class AsposeExamples
		<AssemblyInitialize> _
		Public Shared Sub AssemblyInitialize(ByVal context As Microsoft.VisualStudio.TestTools.UnitTesting.TestContext)
			Main()
		End Sub

		<SetUp> _
		Public Shared Sub AssemblySetup()
			Main()
		End Sub

		<AssemblyCleanup> _
		Public Shared Sub AssemblyCleanup()
			TestHelper.Cleanup()
		End Sub

		Public Shared Sub Main()
		    ' Provides an introduction of how to use this example project.
			TestHelper.ShowIntroForm()
		End Sub
	End Class
End Namespace