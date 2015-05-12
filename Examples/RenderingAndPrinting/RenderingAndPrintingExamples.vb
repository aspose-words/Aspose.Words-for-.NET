Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.RenderingAndPrinting
	<TestClass, TestFixture> _
	Public Class RenderingAndPrintingExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub EnumerateLayoutElements()
	            TestHelper.SetDataDir("RenderingAndPrinting/EnumerateLayoutElements")
	
	            EnumerateLayoutElementsExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub DocumentLayoutHelper()
	            TestHelper.SetDataDir("RenderingAndPrinting/DocumentLayoutHelper")
	
	            DocumentLayoutHelperExample.Program.Main()
        End Sub

	End Class
End Namespace