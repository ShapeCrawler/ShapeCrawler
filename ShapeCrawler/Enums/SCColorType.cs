using System.Diagnostics.CodeAnalysis;

namespace ShapeCrawler
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public enum SCColorType
    {
        NotDefined = 0,
        RGB = 1,
        RGBPercentage = 2,
        HSL = 3,
        Scheme = 4,
        System = 5,
        Preset = 6
    }
}