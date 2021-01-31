namespace ShapeCrawler.Models
{
    public interface IShape
    {
        uint Id { get; }

        long X { get; set; }

        long Y { get; set; }

        long Width { get; set; }

        long Height { get; }

        GeometryType GeometryType { get; }

        Placeholder Placeholder { get; }
    }
}