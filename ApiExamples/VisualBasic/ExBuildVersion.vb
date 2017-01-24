Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExBuildVersion
		Inherits ApiExampleBase
		<Test> _
		Public Sub ShowBuildVersionInfo()
			'ExStart
			'ExFor:BuildVersionInfo
			'ExSummary:Shows how to use BuildVersionInfo to obtain information about this product.
			Console.WriteLine("I am currently using {0}, version number {1}.", BuildVersionInfo.Product, BuildVersionInfo.Version)
			'ExEnd
		End Sub
	End Class
End Namespace
