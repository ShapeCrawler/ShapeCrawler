using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Colors;

internal sealed class PresentationColor
{
    private readonly OpenXmlPart openXmlPart;

    internal PresentationColor(OpenXmlPart openXmlPart)
    {
        this.openXmlPart = openXmlPart;
    }
    
    internal IndentFont? PresentationFontOrThemeFontOrNull(int indentLevel)
    {
        var sdkPresDoc = (PresentationDocument)this.openXmlPart.OpenXmlPackage;
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
        var aColorScheme = GetColorScheme(this.openXmlPart);
        return this.GetColorValue(aColorScheme, aSchemeColorValue);
    }
    
    private static string GetRgbOrSystemColor(A.Color2Type colorType)
    {
        return colorType.RgbColorModelHex != null
            ? colorType.RgbColorModelHex.Val!.Value!
            : colorType.SystemColor!.LastColor!.Value!;
    }

    private static A.ColorScheme GetColorScheme(OpenXmlPart sdkTypedOpenXmlPart)
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
    
    private static string GetColorFromScheme(A.ColorScheme aColorScheme, string fontSchemeColor)
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
    
    private string GetColorValue(A.ColorScheme aColorScheme, A.SchemeColorValues aSchemeColorValue)
    {
        if(aSchemeColorValue == A.SchemeColorValues.Dark1)
        {
            return GetRgbOrSystemColor(aColorScheme.Dark1Color!);
        }

        if (aSchemeColorValue == A.SchemeColorValues.Light1)
        {
            return GetRgbOrSystemColor(aColorScheme.Light1Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Dark2)
        {
            return GetRgbOrSystemColor(aColorScheme.Dark2Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Light2)
        {
            return GetRgbOrSystemColor(aColorScheme.Light2Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent1)
        {
            return GetRgbOrSystemColor(aColorScheme.Accent1Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent2)
        {
            return GetRgbOrSystemColor(aColorScheme.Accent2Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent3)
        {
            return GetRgbOrSystemColor(aColorScheme.Accent3Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent4)
        {
            return GetRgbOrSystemColor(aColorScheme.Accent4Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent5)
        {
            return GetRgbOrSystemColor(aColorScheme.Accent5Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent6)
        {
            return GetRgbOrSystemColor(aColorScheme.Accent6Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Hyperlink)
        {
            return GetRgbOrSystemColor(aColorScheme.Hyperlink!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.FollowedHyperlink)
        {
            return GetRgbOrSystemColor(aColorScheme.FollowedHyperlinkColor!);
        }

        return this.GetThemeMappedColor(aSchemeColorValue);
    }
    
    private string GetThemeMappedColor(A.SchemeColorValues themeColor)
    {
        var pColorMap = this.openXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.ColorMap!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.SlideMaster.ColorMap!,
            _ => ((SlideMasterPart)this.openXmlPart).SlideMaster.ColorMap!
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
        var aColorScheme = GetColorScheme(this.openXmlPart);
        return GetColorFromScheme(aColorScheme, fontSchemeColor);
    }
}