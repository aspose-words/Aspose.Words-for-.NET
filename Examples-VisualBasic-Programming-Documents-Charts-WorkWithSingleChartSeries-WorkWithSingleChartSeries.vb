' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' Get first series.
Dim series0 As ChartSeries = shape.Chart.Series(0)

' Get second series.
Dim series1 As ChartSeries = shape.Chart.Series(1)

' Change first series name.
series0.Name = "My Name1"

' Change second series name.
series1.Name = "My Name2"

' You can also specify whether the line connecting the points on the chart shall be smoothed using Catmull-Rom splines.
series0.Smooth = True
series1.Smooth = True
