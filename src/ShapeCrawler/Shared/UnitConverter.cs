using System;

namespace ShapeCrawler.Shared;

internal static class UnitConverter
{
    private const float HorizontalResolutionDpi = 96;
    private const float VerticalResolutionDpi = 96;
    private const double AngleToDegrees = 1 / 60000d;

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

    internal static long CentimeterToEmu(double centimeter)
    {
        return (long)(centimeter / 0.000002734);
    }

    internal static int CentimeterToPixel(double centimeter)
    {
        return (int)(centimeter * 96 / 2.54);
    }

    internal static double EmuToPoint(int emu)
    {
        return emu * 1.0 / 12700; // 1pt = 12700 EMUs (http://officeopenxml.com/drwSp-outline.php)
    }

    internal static int PointToEmu(double point)
    {
        return (int)(point * 12700); // 1pt = 12700 EMUs (http://officeopenxml.com/drwSp-outline.php)
    }

    internal static float PointToPixel(double points)
    {
        return (float)(points * 96 / 72);
    }

    internal static double AngleValueToDegrees(int angle)
    {
        return angle * AngleToDegrees;
    }
}