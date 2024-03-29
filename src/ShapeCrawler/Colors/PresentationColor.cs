using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Colors;

internal sealed class PresentationColor
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;

    internal PresentationColor(OpenXmlPart sdkTypedOpenXmlPart)
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
        var aColorScheme = this.GetColorScheme(this.sdkTypedOpenXmlPart);
        return this.GetColorValue(aColorScheme, aSchemeColorValue);
    }
    
    private string GetRgbOrSystemColor(A.Color2Type colorType)
    {
        return colorType.RgbColorModelHex != null
            ? colorType.RgbColorModelHex.Val!.Value!
            : colorType.SystemColor!.LastColor!.Value!;
    }
    
    private string GetColorValue(A.ColorScheme aColorScheme, A.SchemeColorValues aSchemeColorValue)
    {
        if(aSchemeColorValue == A.SchemeColorValues.Dark1)
        {
            return this.GetRgbOrSystemColor(aColorScheme.Dark1Color!);
        }

        if (aSchemeColorValue == A.SchemeColorValues.Light1)
        {
            return this.GetRgbOrSystemColor(aColorScheme.Light1Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Dark2)
        {
            return this.GetRgbOrSystemColor(aColorScheme.Dark2Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Light2)
        {
            return this.GetRgbOrSystemColor(aColorScheme.Light2Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent1)
        {
            return this.GetRgbOrSystemColor(aColorScheme.Accent1Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent2)
        {
            return this.GetRgbOrSystemColor(aColorScheme.Accent2Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent3)
        {
            return this.GetRgbOrSystemColor(aColorScheme.Accent3Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent4)
        {
            return this.GetRgbOrSystemColor(aColorScheme.Accent4Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent5)
        {
            return this.GetRgbOrSystemColor(aColorScheme.Accent5Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Accent6)
        {
            return this.GetRgbOrSystemColor(aColorScheme.Accent6Color!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.Hyperlink)
        {
            return this.GetRgbOrSystemColor(aColorScheme.Hyperlink!);
        }
        
        if (aSchemeColorValue == A.SchemeColorValues.FollowedHyperlink)
        {
            return this.GetRgbOrSystemColor(aColorScheme.FollowedHyperlinkColor!);
        }

        return this.GetThemeMappedColor(aSchemeColorValue);
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
        var aColorScheme = this.GetColorScheme(this.sdkTypedOpenXmlPart);
        return this.GetColorFromScheme(aColorScheme, fontSchemeColor);
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