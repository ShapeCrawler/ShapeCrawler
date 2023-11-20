using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Colors;

internal sealed class PresentationColor
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;

    internal PresentationColor(TypedOpenXmlPart sdkTypedOpenXmlPart)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
    }

    #region APIs

    internal IndentFont? PresentationFontOrThemeFontOrNull(int indentLevel)
    {
        var sdkPresDoc = (PresentationDocument)this.sdkTypedOpenXmlPart.OpenXmlPackage;
        var pDefaultTextStyle = sdkPresDoc.PresentationPart!.Presentation.DefaultTextStyle;
        if (pDefaultTextStyle != null)
        {
            var pDefaultTextStyleFont = new IndentFonts(pDefaultTextStyle).FontOrNull(indentLevel);
            if (pDefaultTextStyleFont != null)
            {
                return pDefaultTextStyleFont;
            }
        }

        var aTextDefault = sdkPresDoc.PresentationPart!.ThemePart?.Theme.ObjectDefaults!
            .TextDefault;
        return aTextDefault != null
            ? new IndentFonts(aTextDefault).FontOrNull(indentLevel)
            : null;
    }

    internal string ThemeColorHex(A.SchemeColorValues aSchemeColorValue)
    {
        var aColorScheme = GetColorScheme(this.sdkTypedOpenXmlPart);
        return GetColorValue(aColorScheme, aSchemeColorValue);
    }
    
    private string GetRgbOrSystemColor(A.Color2Type colorType)
    {
        return colorType.RgbColorModelHex != null
            ? colorType.RgbColorModelHex.Val!.Value!
            : colorType.SystemColor!.LastColor!.Value!;
    }
    
    private string GetColorValue(A.ColorScheme aColorScheme, A.SchemeColorValues aSchemeColorValue)
    {
        return aSchemeColorValue switch
        {
            A.SchemeColorValues.Dark1 => GetRgbOrSystemColor(aColorScheme.Dark1Color!),
            A.SchemeColorValues.Light1 => GetRgbOrSystemColor(aColorScheme.Light1Color!),
            A.SchemeColorValues.Dark2 => GetRgbOrSystemColor(aColorScheme.Dark2Color!),
            A.SchemeColorValues.Light2 => GetRgbOrSystemColor(aColorScheme.Light2Color!),
            A.SchemeColorValues.Accent1 => GetRgbOrSystemColor(aColorScheme.Accent1Color!),
            A.SchemeColorValues.Accent2 => GetRgbOrSystemColor(aColorScheme.Accent2Color!),
            A.SchemeColorValues.Accent3 => GetRgbOrSystemColor(aColorScheme.Accent3Color!),
            A.SchemeColorValues.Accent4 => GetRgbOrSystemColor(aColorScheme.Accent4Color!),
            A.SchemeColorValues.Accent5 => GetRgbOrSystemColor(aColorScheme.Accent5Color!),
            A.SchemeColorValues.Accent6 => GetRgbOrSystemColor(aColorScheme.Accent6Color!),
            A.SchemeColorValues.Hyperlink => GetRgbOrSystemColor(aColorScheme.Hyperlink!),
            A.SchemeColorValues.FollowedHyperlink => GetRgbOrSystemColor(aColorScheme.FollowedHyperlinkColor!),
            _ => this.GetThemeMappedColor(aSchemeColorValue)
        };
    }
    
    private A.ColorScheme GetColorScheme(OpenXmlPart sdkTypedOpenXmlPart)
    {
        return sdkTypedOpenXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            _ => ((SlideMasterPart)sdkTypedOpenXmlPart).ThemePart!.Theme.ThemeElements!.ColorScheme!
        };
    }
    
    #endregion APIs

    private string GetThemeMappedColor(A.SchemeColorValues themeColor)
    {
        var pColorMap = this.sdkTypedOpenXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.ColorMap!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.SlideMaster.ColorMap!,
            _ => ((SlideMasterPart)this.sdkTypedOpenXmlPart).SlideMaster.ColorMap!
        };
        if (themeColor == A.SchemeColorValues.Text1)
        {
            return this.GetThemeColorByString(pColorMap.Text1!.ToString() !);
        }

        if (themeColor == A.SchemeColorValues.Text2)
        {
            return this.GetThemeColorByString(pColorMap.Text2!.ToString() !);
        }

        if (themeColor == A.SchemeColorValues.Background1)
        {
            return this.GetThemeColorByString(pColorMap.Background1!.ToString() !);
        }

        return this.GetThemeColorByString(pColorMap.Background2!.ToString() !);
    }

    private string GetThemeColorByString(string fontSchemeColor)
    {
        var aColorScheme = GetColorScheme(this.sdkTypedOpenXmlPart);
        return GetColorFromScheme(aColorScheme, fontSchemeColor);
    }
    
    private string GetColorFromScheme(A.ColorScheme aColorScheme, string fontSchemeColor)
    {
        var colorMap = new Dictionary<string, Func<A.Color2Type>>
        {
            ["dk1"] = () => aColorScheme.Dark1Color!,
            ["lt1"] = () => aColorScheme.Light1Color!,
            ["dk2"] = () => aColorScheme.Dark2Color!,
            ["lt2"] = () => aColorScheme.Light2Color!,
            ["accent1"] = () => aColorScheme.Accent1Color!,
            ["accent2"] = () => aColorScheme.Accent2Color!,
            ["accent3"] = () => aColorScheme.Accent3Color!,
            ["accent4"] = () => aColorScheme.Accent4Color!,
            ["accent5"] = () => aColorScheme.Accent5Color!,
            ["accent6"] = () => aColorScheme.Accent6Color!,
            ["hyperlink"] = () => aColorScheme.Hyperlink!
        };

        if (colorMap.TryGetValue(fontSchemeColor, out var getColor))
        {
            var colorType = getColor();
            return colorType.RgbColorModelHex != null
                ? colorType.RgbColorModelHex.Val!.Value!
                : colorType.SystemColor!.LastColor!.Value!;
        }

        // Default or fallback color
        return aColorScheme.Hyperlink!.RgbColorModelHex != null
            ? aColorScheme.Hyperlink.RgbColorModelHex.Val!.Value!
            : aColorScheme.Hyperlink.SystemColor!.LastColor!.Value!;
    }
}