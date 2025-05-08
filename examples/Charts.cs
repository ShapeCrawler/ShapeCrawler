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
    
    [Test, Explicit]
    public void Add_Scatter_chart()
    {
        var pres = new Presentation();
        var shapes = pres.Slide(1).Shapes;
        int x = 100;
        int y = 100;
        int width = 500;
        int height = 300;
        var pointValues = new Dictionary<double, double>
        {
            { 1.0, 5.2 },
            { 2.0, 7.3 },
            { 3.0, 8.1 },
            { 4.0, 9.5 },
            { 5.0, 12.3 }
        };
        string seriesName = "Data Series";
        
        shapes.AddScatterChart(x, y, width, height, pointValues, seriesName);
    }

    [Test, Explicit]
    public void Add_Stacked_Column_chart()
    {
        var pres = new Presentation();
        var shapes = pres.Slide(1).Shapes;
        int x = 100;
        int y = 100;
        int width = 500;
        int height = 300;
        var categoryValues = new Dictionary<string, IList<double>>
        {
            { "Category 1", new List<double> { 10, 20 } },
            { "Category 2", new List<double> { 30, 40 } },
            { "Category 3", new List<double> { 50, 60 } }
        };
        var seriesNames = new List<string> { "Series 1", "Series 2" };

        shapes.AddStackedColumnChart(x, y, width, height, categoryValues, seriesNames);
    }
}