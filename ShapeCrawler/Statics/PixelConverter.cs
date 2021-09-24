using System.Drawing;

namespace ShapeCrawler.Statics
{
    internal class PixelConverter
    {
        private static readonly float HorizontalResolution;
        private static readonly float VerticalResolution;

        static PixelConverter()
        {
            var bm = new Bitmap(100, 100);
            HorizontalResolution = bm.HorizontalResolution;
            VerticalResolution = bm.VerticalResolution;
        }

        internal static int HorizontalEmuToPixel(long horizontalEmu)
        {
            return (int)(horizontalEmu * HorizontalResolution / 914400);
        }

        internal static int VerticalEmuToPixel(long verticalEmu)
        {
            return (int)(verticalEmu * VerticalResolution / 914400);
        }
    }
}