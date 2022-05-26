using ShapeCrawler;

internal class ChartSample
{
    internal void UpdateSeriesValue()
    {
        var pres = SCPresentation.Open("test.pptx", true);
        var chart = pres.Slides[0].Shapes.GetByName<IChart>("Chart 1");
        var point = chart.SeriesCollection[0].Points[0];
        point.Value = 10;
    }
}