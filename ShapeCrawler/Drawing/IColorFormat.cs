using System.Drawing;

namespace ShapeCrawler.Drawing
{
    public interface IColorFormat
    {
        SCColorType ColorType { get; }

#if DEBUG
        Color Color { get; set; }
#else
        Color Color { get; }
#endif

    }
}