using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Wrappers;

internal sealed record SDKSlidePartWrap
{
    private readonly SlidePart sdkSlidePart;

    internal SDKSlidePartWrap(SlidePart sdkSlidePart)
    {
        this.sdkSlidePart = sdkSlidePart;
    }

    internal PresentationDocument SDKPresentationDocument()
    {
        return (PresentationDocument)this.sdkSlidePart.OpenXmlPackage;
    }

    internal ParagraphLevelFont? PresentationFontOrThemeFontOrNull(int paraLevel)
    {
        var pDefaultTextStyle = this.SDKPresentationDocument().PresentationPart!.Presentation.DefaultTextStyle;
        if (pDefaultTextStyle != null)
        {
            var pDefaultTextStyleFont = new ParagraphLevelFonts(pDefaultTextStyle).FontOrNull(paraLevel);
            if (pDefaultTextStyleFont != null)
            {
                return pDefaultTextStyleFont;
            }
        }

        var aTextDefault = this.SDKPresentationDocument().PresentationPart!.ThemePart?.Theme.ObjectDefaults!
            .TextDefault;
        return aTextDefault != null
            ? new ParagraphLevelFonts(aTextDefault).FontOrNull(paraLevel)
            : null;
    }

    public string ThemeHexColor(A.SchemeColorValues themeColor)
    {
        var aColorScheme = this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
            .ColorScheme!;
        return themeColor switch
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
            _ => this.GetThemeMappedColor(themeColor)
        };
    }

    internal string GetThemeMappedColor(A.SchemeColorValues themeColor)
    {
        var pColorMap = this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.ColorMap!;
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
        var aColorScheme = this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!.ColorScheme!;
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