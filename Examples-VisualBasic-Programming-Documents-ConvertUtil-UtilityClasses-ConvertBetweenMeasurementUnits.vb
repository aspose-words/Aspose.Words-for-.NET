' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)

Dim pageSetup As PageSetup = builder.PageSetup
pageSetup.TopMargin = ConvertUtil.InchToPoint(1.0)
pageSetup.BottomMargin = ConvertUtil.InchToPoint(1.0)
pageSetup.LeftMargin = ConvertUtil.InchToPoint(1.5)
pageSetup.RightMargin = ConvertUtil.InchToPoint(1.5)
pageSetup.HeaderDistance = ConvertUtil.InchToPoint(0.2)
pageSetup.FooterDistance = ConvertUtil.InchToPoint(0.2)
