using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Colors;

internal sealed class PresentationColor(OpenXmlPart openXmlPart)
{
    internal IndentFont? PresentationOrThemeFontOrNull(int indentLevel)
    {
        var presDocument = (PresentationDocument)openXmlPart.OpenXmlPackage;
        var pDefaultTextStyle = presDocument.PresentationPart!.Presentation.DefaultTextStyle;
        if (pDefaultTextStyle != null)
        {
            var pDefaultTextStyleFont = new IndentFonts(pDefaultTextStyle).FontOrNull(indentLevel);
            if (pDefaultTextStyleFont != null)
            {
                return pDefaultTextStyleFont;
            }
        }

        var aTextDefault = presDocument.PresentationPart!.ThemePart?.Theme.ObjectDefaults!
            .TextDefault;
        return aTextDefault != null
            ? new IndentFonts(aTextDefault).FontOrNull(indentLevel)
            : null;
    }

    internal string ThemeColorHex(A.SchemeColorValues aSchemeColorValue)
    {
        var aColorScheme = GetColorScheme(openXmlPart);

        return this.GetColorValue(aColorScheme, aSchemeColorValue);
    }

    private static string GetRgbOrSystemColor(A.Color2Type colorType)
    {
        return colorType.RgbColorModelHex != null
            ? colorType.RgbColorModelHex.Val!.Value!
            : colorType.SystemColor!.LastColor!.Value!;
    }

    private static A.ColorScheme GetColorScheme(OpenXmlPart openXmlPart)
    {
        return openXmlPart switch
        {
            SlidePart slidePart => slidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            SlideLayoutPart slideLayoutPart => slideLayoutPart.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            NotesSlidePart notesSlidePart =>
                notesSlidePart.GetParentParts().OfType<SlidePart>().First().SlideLayoutPart!.SlideMasterPart!.ThemePart!
                    .Theme.ThemeElements!
                    .ColorScheme!,
            _ => ((SlideMasterPart)openXmlPart).ThemePart!.Theme.ThemeElements!.ColorScheme!
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
        var mapping = new Dictionary<A.SchemeColorValues, Func<A.Color2Type>>
        {
            { A.SchemeColorValues.Dark1, () => aColorScheme.Dark1Color! },
            { A.SchemeColorValues.Light1, () => aColorScheme.Light1Color! },
            { A.SchemeColorValues.Dark2, () => aColorScheme.Dark2Color! },
            { A.SchemeColorValues.Light2, () => aColorScheme.Light2Color! },
            { A.SchemeColorValues.Accent1, () => aColorScheme.Accent1Color! },
            { A.SchemeColorValues.Accent2, () => aColorScheme.Accent2Color! },
            { A.SchemeColorValues.Accent3, () => aColorScheme.Accent3Color! },
            { A.SchemeColorValues.Accent4, () => aColorScheme.Accent4Color! },
            { A.SchemeColorValues.Accent5, () => aColorScheme.Accent5Color! },
            { A.SchemeColorValues.Accent6, () => aColorScheme.Accent6Color! },
            { A.SchemeColorValues.Hyperlink, () => aColorScheme.Hyperlink! },
            { A.SchemeColorValues.FollowedHyperlink, () => aColorScheme.FollowedHyperlinkColor! }
        };

        if (mapping.TryGetValue(aSchemeColorValue, out var getColorType))
        {
            return GetRgbOrSystemColor(getColorType());
        }

        return this.GetThemeMappedColor(aSchemeColorValue);
    }

    private string GetThemeMappedColor(A.SchemeColorValues themeColor)
    {
        var pColorMap = openXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.ColorMap!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.SlideMaster.ColorMap!,
            NotesSlidePart notesSlidePart =>
                notesSlidePart.GetParentParts().OfType<SlidePart>().First().SlideLayoutPart!.SlideMasterPart!
                    .SlideMaster.ColorMap!,
            _ => ((SlideMasterPart)openXmlPart).SlideMaster.ColorMap!
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
        var aColorScheme = GetColorScheme(openXmlPart);
        return GetColorFromScheme(aColorScheme, fontSchemeColor);
    }
}