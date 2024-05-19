using System;

namespace ShapeCrawler.Shared;

internal static class UnitConverter
{
    private const int HorizontalResolutionDpi = 96;
    private const int VerticalResolutionDpi = 96;
    private const double AngleToDegrees = 1 / 60000d;
    private const int EmusPerInch = 914400;
    private const int EmusPerCentimeter = 360000;
    private const int EmusPerPoint = 12700;

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
        return (long)Math.Round(horizontalPixels * EmusPerInch / (decimal)HorizontalResolutionDpi);
    }

    internal static long VerticalPixelToEmu(decimal verticalPixels)
    {
        return (long)Math.Round(verticalPixels * EmusPerInch / (decimal)VerticalResolutionDpi);
    }

    internal static decimal EmuToCentimeter(long emu)
    {
        return emu / (decimal)EmusPerCentimeter;
    }

    internal static long CentimeterToEmu(decimal centimeter)
    {
        return (long)Math.Round(centimeter * EmusPerCentimeter);
    }

    internal static decimal CentimeterToPixel(decimal centimeter)
    {
        return HorizontalEmuToPixel(CentimeterToEmu(centimeter));
    }

    internal static decimal EmuToPoint(long emu)
    {
        return emu / (decimal)EmusPerPoint; // 1pt = 12700 EMUs (http://officeopenxml.com/drwSp-outline.php)
    }

    internal static long PointToEmu(decimal point)
    {
        return (long)Math.Round(point * EmusPerPoint); // 1pt = 12700 EMUs (http://officeopenxml.com/drwSp-outline.php)
    }

    internal static decimal PointToPixel(decimal point)
    {
        return HorizontalEmuToPixel(PointToEmu(point));
    }

    internal static double AngleValueToDegrees(int angle)
    {
        return angle * AngleToDegrees;
    }
}