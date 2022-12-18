using DocumentFormat.OpenXml;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal static class HexParser
{
    internal static (SCColorType, string) FromSolidFill(TypedOpenXmlCompositeElement typedElement, SCSlideMaster slideMaster)
    {
        var colorHexVariant = GetWithoutScheme(typedElement);
        if (colorHexVariant is not null)
        {
            return ((SCColorType, string))colorHexVariant;
        }

        var aSchemeColor = typedElement.GetFirstChild<A.SchemeColor>() !;
        var fromScheme = GetByThemeColorScheme(aSchemeColor.Val!, slideMaster); 
        return (SCColorType.Scheme, fromScheme);
    }

    internal static (SCColorType, string)? GetWithoutScheme(TypedOpenXmlCompositeElement typedElement)
    {
        var aSrgbClr = typedElement.GetFirstChild<A.RgbColorModelHex>();
        string colorHexVariant;
        if (aSrgbClr != null)
        {
            colorHexVariant = aSrgbClr.Val!;
            {
                return (SCColorType.RGB, colorHexVariant);
            }
        }

        var aSysClr = typedElement.GetFirstChild<A.SystemColor>();
        if (aSysClr != null)
        {
            colorHexVariant = aSysClr.LastColor!;
            {
                return (SCColorType.System, colorHexVariant);
            }
        }

        var aPresetColor = typedElement.GetFirstChild<A.PresetColor>();
        if (aPresetColor != null)
        {
            var coloName = aPresetColor.Val!.Value.ToString();
            {
                return (SCColorType.Preset, SCColorTranslator.HexFromName(coloName));
            }
        }

        return null;
    }

    private static string GetByThemeColorScheme(A.SchemeColorValues fontSchemeColor, SCSlideMaster slideMaster)
    {
        var themeAColorScheme = slideMaster.ThemePart.Theme.ThemeElements!.ColorScheme!;
        return fontSchemeColor switch
        {
            A.SchemeColorValues.Dark1 => themeAColorScheme.Dark1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Dark1Color.RgbColorModelHex!.Val!.Value!
                : themeAColorScheme.Dark1Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Light1 => themeAColorScheme.Light1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Light1Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Light1Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Dark2 => themeAColorScheme.Dark2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Dark2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Dark2Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Light2 => themeAColorScheme.Light2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Light2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Light2Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent1 => themeAColorScheme.Accent1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent1Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent1Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent2 => themeAColorScheme.Accent2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent2Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent3 => themeAColorScheme.Accent3Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent3Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent3Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent4 => themeAColorScheme.Accent4Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent4Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent4Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent5 => themeAColorScheme.Accent5Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent5Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent5Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent6 => themeAColorScheme.Accent6Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent6Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent6Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Hyperlink => themeAColorScheme.Hyperlink!.RgbColorModelHex != null
                ? themeAColorScheme.Hyperlink.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Hyperlink.SystemColor!.LastColor!.Value!,
            _ => GetThemeMappedColor(fontSchemeColor, slideMaster)
        };
    }

    private static string GetThemeMappedColor(A.SchemeColorValues fontSchemeColor, SCSlideMaster slideMaster)
    {
        var slideMasterPColorMap = slideMaster.PSlideMaster.ColorMap;
        if (fontSchemeColor == A.SchemeColorValues.Text1)
        {
            return GetThemeColorByString(slideMasterPColorMap!.Text1!.ToString() !, slideMaster);
        }

        if (fontSchemeColor == A.SchemeColorValues.Text2)
        {
            return GetThemeColorByString(slideMasterPColorMap!.Text2!.ToString() !, slideMaster);
        }

        if (fontSchemeColor == A.SchemeColorValues.Background1)
        {
            return GetThemeColorByString(slideMasterPColorMap!.Background1!.ToString() !, slideMaster);
        }

        return GetThemeColorByString(slideMasterPColorMap!.Background2!.ToString() !, slideMaster);
    }

    private static string GetThemeColorByString(string fontSchemeColor, SCSlideMaster slideMaster)
    {
        var themeAColorScheme = slideMaster.ThemePart.Theme.ThemeElements!.ColorScheme!;
        return fontSchemeColor switch
        {
            "dk1" => themeAColorScheme.Dark1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Dark1Color.RgbColorModelHex!.Val!.Value!
                : themeAColorScheme.Dark1Color.SystemColor!.LastColor!.Value!,
            "lt1" => themeAColorScheme.Light1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Light1Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Light1Color.SystemColor!.LastColor!.Value!,
            "dk2" => themeAColorScheme.Dark2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Dark2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Dark2Color.SystemColor!.LastColor!.Value!,
            "lt2" => themeAColorScheme.Light2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Light2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Light2Color.SystemColor!.LastColor!.Value!,
            "accent1" => themeAColorScheme.Accent1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent1Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent1Color.SystemColor!.LastColor!.Value!,
            "accent2" => themeAColorScheme.Accent2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent2Color.SystemColor!.LastColor!.Value!,
            "accent3" => themeAColorScheme.Accent3Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent3Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent3Color.SystemColor!.LastColor!.Value!,
            "accent4" => themeAColorScheme.Accent4Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent4Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent4Color.SystemColor!.LastColor!.Value!,
            "accent5" => themeAColorScheme.Accent5Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent5Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent5Color.SystemColor!.LastColor!.Value!,
            "accent6" => themeAColorScheme.Accent6Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent6Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent6Color.SystemColor!.LastColor!.Value!,
            _ => themeAColorScheme.Hyperlink!.RgbColorModelHex != null
                ? themeAColorScheme.Hyperlink.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Hyperlink.SystemColor!.LastColor!.Value!
        };
    }
}