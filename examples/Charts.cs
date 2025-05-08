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
    
    [Test, Explicit]
    public static void Update_chart_category()
    {
        using var pres = new Presentation("pres.pptx");
        var slide = pres.Slide(1);
        var chart = slide.First<IChart>();

        if (chart.Type == ChartType.BarChart)
        {
            Console.WriteLine("Chart type is BarChart");
        }
        
        chart.Categories![0].Name = "Price";
    }
}