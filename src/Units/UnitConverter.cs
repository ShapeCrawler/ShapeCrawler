using System;

namespace ShapeCrawler.Units;

internal static class UnitConverter
{
    private const int HorizontalResolutionDpi = 96;
    private const int VerticalResolutionDpi = 96;
    private const int EmusPerInch = 914400;

    internal static long HorizontalPixelToEmu(decimal horizontalPixels) =>
        (long)Math.Round(horizontalPixels * EmusPerInch / HorizontalResolutionDpi);

    internal static long VerticalPixelToEmu(decimal verticalPixels) =>
        (long)Math.Round(verticalPixels * EmusPerInch / VerticalResolutionDpi);
}