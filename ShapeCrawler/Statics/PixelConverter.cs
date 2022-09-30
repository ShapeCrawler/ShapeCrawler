using System.Diagnostics.CodeAnalysis;

namespace ShapeCrawler.Statics
{
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600:Elements should be documented", Justification = "Will be converted to internal")]
    internal static class PixelConverter
    {
        private static readonly float horizontalResolution;
        private static readonly float verticalResolution;

        static PixelConverter()
        {
            horizontalResolution = 96;
            verticalResolution = 96;
        }

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