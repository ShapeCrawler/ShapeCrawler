using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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
        var aColorScheme = this.sdkTypedOpenXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            _ => ((SlideMasterPart)this.sdkTypedOpenXmlPart).ThemePart!.Theme.ThemeElements!.ColorScheme!
        };
        return aSchemeColorValue switch
        {
            A.SchemeColorValues.Dark1 => aColorScheme.Dark1Color!.RgbColorModelHex != null
                ? aColorScheme.Dark1Color.RgbColorModelHex!.Val!.Value!
                : aColorScheme.Dark1Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Light1 => aColorScheme.Light1Color!.RgbColorModelHex != null
                ? aColorScheme.Light1Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Light1Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Dark2 => aColorScheme.Dark2Color!.RgbColorModelHex != null
                ? aColorScheme.Dark2Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Dark2Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Light2 => aColorScheme.Light2Color!.RgbColorModelHex != null
                ? aColorScheme.Light2Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Light2Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent1 => aColorScheme.Accent1Color!.RgbColorModelHex != null
                ? aColorScheme.Accent1Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent1Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent2 => aColorScheme.Accent2Color!.RgbColorModelHex != null
                ? aColorScheme.Accent2Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent2Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent3 => aColorScheme.Accent3Color!.RgbColorModelHex != null
                ? aColorScheme.Accent3Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent3Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent4 => aColorScheme.Accent4Color!.RgbColorModelHex != null
                ? aColorScheme.Accent4Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent4Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent5 => aColorScheme.Accent5Color!.RgbColorModelHex != null
                ? aColorScheme.Accent5Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent5Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent6 => aColorScheme.Accent6Color!.RgbColorModelHex != null
                ? aColorScheme.Accent6Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent6Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Hyperlink => aColorScheme.Hyperlink!.RgbColorModelHex != null
                ? aColorScheme.Hyperlink.RgbColorModelHex.Val!.Value!
                : aColorScheme.Hyperlink.SystemColor!.LastColor!.Value!,
            _ => this.GetThemeMappedColor(aSchemeColorValue)
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
        var aColorScheme = this.sdkTypedOpenXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            _ => ((SlideMasterPart)this.sdkTypedOpenXmlPart).ThemePart!.Theme.ThemeElements!.ColorScheme!
        };
        return fontSchemeColor switch
        {
            "dk1" => aColorScheme.Dark1Color!.RgbColorModelHex != null
                ? aColorScheme.Dark1Color.RgbColorModelHex!.Val!.Value!
                : aColorScheme.Dark1Color.SystemColor!.LastColor!.Value!,
            "lt1" => aColorScheme.Light1Color!.RgbColorModelHex != null
                ? aColorScheme.Light1Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Light1Color.SystemColor!.LastColor!.Value!,
            "dk2" => aColorScheme.Dark2Color!.RgbColorModelHex != null
                ? aColorScheme.Dark2Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Dark2Color.SystemColor!.LastColor!.Value!,
            "lt2" => aColorScheme.Light2Color!.RgbColorModelHex != null
                ? aColorScheme.Light2Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Light2Color.SystemColor!.LastColor!.Value!,
            "accent1" => aColorScheme.Accent1Color!.RgbColorModelHex != null
                ? aColorScheme.Accent1Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent1Color.SystemColor!.LastColor!.Value!,
            "accent2" => aColorScheme.Accent2Color!.RgbColorModelHex != null
                ? aColorScheme.Accent2Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent2Color.SystemColor!.LastColor!.Value!,
            "accent3" => aColorScheme.Accent3Color!.RgbColorModelHex != null
                ? aColorScheme.Accent3Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent3Color.SystemColor!.LastColor!.Value!,
            "accent4" => aColorScheme.Accent4Color!.RgbColorModelHex != null
                ? aColorScheme.Accent4Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent4Color.SystemColor!.LastColor!.Value!,
            "accent5" => aColorScheme.Accent5Color!.RgbColorModelHex != null
                ? aColorScheme.Accent5Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent5Color.SystemColor!.LastColor!.Value!,
            "accent6" => aColorScheme.Accent6Color!.RgbColorModelHex != null
                ? aColorScheme.Accent6Color.RgbColorModelHex.Val!.Value!
                : aColorScheme.Accent6Color.SystemColor!.LastColor!.Value!,
            _ => aColorScheme.Hyperlink!.RgbColorModelHex != null
                ? aColorScheme.Hyperlink.RgbColorModelHex.Val!.Value!
                : aColorScheme.Hyperlink.SystemColor!.LastColor!.Value!
        };
    }
}