Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.RenderingAndPrinting.RenderingToImage
	<TestClass, TestFixture> _
	Public Class RenderingToImageExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub SaveAsMultipageTiff()
	            TestHelper.SetDataDir("RenderingAndPrinting/RenderingToImage/SaveAsMultipageTiff")
	
	            SaveAsMultipageTiffExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub RenderShapes()
	            TestHelper.SetDataDir("RenderingAndPrinting/RenderingToImage/RenderShapes")
	
	            RenderShapesExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub ImageColorFilters()
	            TestHelper.SetDataDir("RenderingAndPrinting/RenderingToImage/ImageColorFilters")
	
	            ImageColorFiltersExample.Program.Main()
        End Sub

	End Class
End Namespace