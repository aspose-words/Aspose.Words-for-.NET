Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.MailMergeAndReporting
	<TestClass, TestFixture> _
	Public Class MailMergeAndReportingExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub ApplyCustomLogicToEmptyRegions()
	            TestHelper.SetDataDir("MailMergeAndReporting/ApplyCustomLogicToEmptyRegions")
	
	            ApplyCustomLogicToEmptyRegionsExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub XMLMailMerge()
	            TestHelper.SetDataDir("MailMergeAndReporting/XMLMailMerge")
	
	            XMLMailMergeExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub RemoveEmptyRegions()
	            TestHelper.SetDataDir("MailMergeAndReporting/RemoveEmptyRegions")
	
	            RemoveEmptyRegionsExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub NestedMailMerge()
	            TestHelper.SetDataDir("MailMergeAndReporting/NestedMailMerge")
	
	            NestedMailMergeExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub MultipleDocsInMailMerge()
	            TestHelper.SetDataDir("MailMergeAndReporting/MultipleDocsInMailMerge")
	
	            MultipleDocsInMailMergeExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub MailMergeFormFields()
	            TestHelper.SetDataDir("MailMergeAndReporting/MailMergeFormFields")
	
	            MailMergeFormFieldsExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub LINQtoXMLMailMerge()
	            TestHelper.SetDataDir("MailMergeAndReporting/LINQtoXMLMailMerge")
	
	            LINQtoXMLMailMergeExample.Program.Main()
        End Sub

	End Class
End Namespace