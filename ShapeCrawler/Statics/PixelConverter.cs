using System;
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

        internal static int HorizontalEmuToPixel(long horizontalEmus)
        {
            Console.WriteLine($"HorizontalResolution: {HorizontalResolution}");
            return (int)(horizontalEmus * HorizontalResolution / 914400);
        }

        internal static int VerticalEmuToPixel(long verticalEmus)
        {
            return (int)(verticalEmus * VerticalResolution / 914400);
        }

        internal static long HorizontalPixelToEmu(int horizontalPixels)
        {
            return (long)(horizontalPixels * 914400 / HorizontalResolution);
        }

        internal static long VerticalPixelToEmu(int verticalPixels)
        {
            return (long)(verticalPixels * 914400 / VerticalResolution);
        }
    }
}