using System;

namespace ShapeCrawler.Statics;

internal static class UnitConverter
{
    private const float HorizontalResolutionDpi = 96;
    private const float VerticalResolutionDpi = 96;

    internal static int HorizontalEmuToPixel(long horizontalEmus)
    {
        return (int)(horizontalEmus * HorizontalResolutionDpi / 914400);
    }

    internal static int VerticalEmuToPixel(long verticalEmus)
    {
        return (int)(verticalEmus * VerticalResolutionDpi / 914400);
    }

    internal static long HorizontalPixelToEmu(int horizontalPixels)
    {
        return (long)(horizontalPixels * 914400 / HorizontalResolutionDpi);
    }

    internal static long VerticalPixelToEmu(long verticalPixels)
    {
        return (long)(verticalPixels * 914400 / VerticalResolutionDpi);
    }

    internal static double EmuToCentimeter(int emu)
    {
        return Math.Round(emu * 0.000002734, 2);
    }

    internal static int CentimeterToPixel(double centimeter)
    {
        return (int)(centimeter * 37.795275591);
    }
}