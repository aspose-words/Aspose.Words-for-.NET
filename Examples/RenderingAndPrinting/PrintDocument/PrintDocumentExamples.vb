Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.RenderingAndPrinting.PrintDocument
	<TestClass, TestFixture> _
	Public Class PrintDocumentExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub XpsPrint()
	            TestHelper.SetDataDir("RenderingAndPrinting/PrintDocument/XpsPrint")
	
	            XpsPrintExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub MultiplePagesOnSheet()
	            TestHelper.SetDataDir("RenderingAndPrinting/PrintDocument/MultiplePagesOnSheet")
	
	            MultiplePagesOnSheetExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub DocumentPreviewAndPrint()
	            TestHelper.SetDataDir("RenderingAndPrinting/PrintDocument/DocumentPreviewAndPrint")
	
	            DocumentPreviewAndPrintExample.Program.Main()
        End Sub

	End Class
End Namespace