namespace ShapeCrawler;

public interface IXAxis
{
    double[] Values { get; }
    int Minimum { get; set; }
    int Maximum { get; set; }
}