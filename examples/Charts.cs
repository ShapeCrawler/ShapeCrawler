namespace ShapeCrawler.Examples;

public class Charts
{
    [Test, Explicit]
    public void Update_series()
    {
        using var pres = new Presentation("hello world.pptx");
        var chart = pres.Slides[0].Shapes.Shape<IChart>("Chart 1");
        var point = chart.SeriesCollection[0].Points[0];
        point.Value = 10;
    }
}