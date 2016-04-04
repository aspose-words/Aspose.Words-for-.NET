// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// Get first series.
ChartSeries series0 = shape.Chart.Series[0];

// Get second series.
ChartSeries series1 = shape.Chart.Series[1];

// Change first series name.
series0.Name = "My Name1";

// Change second series name.
series1.Name = "My Name2";

// You can also specify whether the line connecting the points on the chart shall be smoothed using Catmull-Rom splines.
series0.Smooth = true;
series1.Smooth = true;
