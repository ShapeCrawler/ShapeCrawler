using ShapeCrawler.Drawing;

namespace ShapeCrawler.AutoShapes
{
    internal class ColorFormat : IColorFormat
    {
        private readonly SCFont _font;

        public ColorFormat(SCFont font)
        {
            this._font = font;
        }

        public SCColorType ColorType { get; set; }
    }
}