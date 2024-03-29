using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Colors;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal static class HexParser
{
    internal static (ColorType, string?) FromSolidFill(A.SolidFill aSolidFill, P.SlideMaster pSlideMaster)
    {
        var colorHexVariant = GetWithoutScheme(aSolidFill);
        if (colorHexVariant is not null)
        {
            return ((ColorType, string))colorHexVariant;
        }

        var aSchemeColor = aSolidFill.GetFirstChild<A.SchemeColor>() !;
        var fromScheme = GetByThemeColorScheme(aSchemeColor.Val!, pSlideMaster); 
        return (ColorType.Theme, fromScheme);
    }

    internal static (ColorType, string)? GetWithoutScheme(OpenXmlCompositeElement typedElement)
    {
        var aSrgbClr = typedElement.GetFirstChild<A.RgbColorModelHex>();
        string colorHexVariant;
        if (aSrgbClr != null)
        {
            colorHexVariant = aSrgbClr.Val!;
            {
                return (ColorType.RGB, colorHexVariant);
            }
        }

        var aSysClr = typedElement.GetFirstChild<A.SystemColor>();
        if (aSysClr != null)
        {
            colorHexVariant = aSysClr.LastColor!;
            {
                return (ColorType.Standard, colorHexVariant);
            }
        }

        var aPresetColor = typedElement.GetFirstChild<A.PresetColor>();
        if (aPresetColor != null)
        {
            var coloName = aPresetColor.Val!.ToString() !;
            {
                return (ColorType.Preset, ColorTranslator.HexFromName(coloName));
            }
        }

        return null;
    }

    private static string? GetByThemeColorScheme(string schemeColor, P.SlideMaster pSlideMaster)
    {
        var hex = GetThemeColorByString(schemeColor, pSlideMaster);

        if (hex == null)
        {
            hex = GetThemeMappedColor(schemeColor, pSlideMaster);
        }

        return hex ?? null;
    }

    private static string? GetThemeMappedColor(string fontSchemeColor, P.SlideMaster pSlideMaster)
    {
        var slideMasterPColorMap = pSlideMaster.ColorMap;
        var targetSchemeColor = slideMasterPColorMap?.GetAttributes().FirstOrDefault(a => a.LocalName == fontSchemeColor);
        return GetThemeColorByString(targetSchemeColor?.Value !, pSlideMaster);
    }

    private static string? GetThemeColorByString(string schemeColor, P.SlideMaster pSlideMaster)
    {
        var themeAColorScheme = pSlideMaster.SlideMasterPart!.ThemePart!.Theme.ThemeElements!.ColorScheme!;
        var color = themeAColorScheme.Elements<A.Color2Type>().FirstOrDefault(c => c.LocalName == schemeColor);
        var hex = color?.RgbColorModelHex?.Val?.Value ?? color?.SystemColor?.LastColor?.Value;
        return hex;
    }
}