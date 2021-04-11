using System.Drawing;

namespace ShapeCrawler.Drawing
{
    public interface IColorFormat
    {
        SCColorType ColorType { get; }
        Color Color { get; }
    }
}