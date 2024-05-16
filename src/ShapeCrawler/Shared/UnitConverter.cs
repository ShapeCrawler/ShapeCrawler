using System;

namespace ShapeCrawler.Shared;

internal static class UnitConverter
{
    private const int HorizontalResolutionDpi = 96;
    private const int VerticalResolutionDpi = 96;
    private const double AngleToDegrees = 1 / 60000d;
    private const int EmusPerInch = 914400;

    internal static decimal HorizontalEmuToPixel(long horizontalEmus)
    {
        return horizontalEmus * HorizontalResolutionDpi / (decimal)EmusPerInch;
    }

    internal static decimal VerticalEmuToPixel(long verticalEmus)
    {
        return verticalEmus * VerticalResolutionDpi / (decimal)EmusPerInch;
    }

    internal static long HorizontalPixelToEmu(decimal horizontalPixels)
    {
        return (long)(horizontalPixels * EmusPerInch / HorizontalResolutionDpi);
    }

    internal static long VerticalPixelToEmu(decimal verticalPixels)
    {
        return (long)(verticalPixels * EmusPerInch / VerticalResolutionDpi);
    }

    internal static double EmuToCentimeter(long emu)
    {
        return Math.Round(emu * 0.000002734, 2);
    }

    internal static decimal CentimeterToEmu(double centimeter)
    {
        return (decimal)(centimeter / 0.000002734);
    }

    internal static decimal CentimeterToPixel(double centimeter)
    {
        return (decimal)(centimeter * 96 / 2.54);
    }

    internal static double EmuToPoint(long emu)
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