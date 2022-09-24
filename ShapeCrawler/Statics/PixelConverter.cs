using System.Drawing;

namespace ShapeCrawler.Statics
{
    public static class PixelConverter // TODO: make internal
    {
        private static float horizontalResolution;
        private static float verticalResolution;

        static PixelConverter()
        {
            var bm = new Bitmap(1, 1);
            if (bm.HorizontalResolution == 0)
            {
                // Set default resolution
                bm.SetResolution(96, 96);
            }

            horizontalResolution = bm.HorizontalResolution;
            verticalResolution = bm.VerticalResolution;
        }

#if TEST

        public static void SetDpi(int dpi)
        {
            horizontalResolution = dpi;
            verticalResolution = dpi;
        }

#endif

        internal static int HorizontalEmuToPixel(long horizontalEmus)
        {
            return (int)(horizontalEmus * horizontalResolution / 914400);
        }

        internal static int VerticalEmuToPixel(long verticalEmus)
        {
            return (int)(verticalEmus * verticalResolution / 914400);
        }

        internal static long HorizontalPixelToEmu(int horizontalPixels)
        {
            return (long)(horizontalPixels * 914400 / horizontalResolution);
        }

        internal static long VerticalPixelToEmu(long verticalPixels)
        {
            return (long)(verticalPixels * 914400 / verticalResolution);
        }
    }
}