using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Wrappers;

internal sealed record SdkSlidePart
{
    private readonly SlidePart sdkSlidePart;

    internal SdkSlidePart(SlidePart sdkSlidePart)
    {
        this.sdkSlidePart = sdkSlidePart;
    }

    #region APIs

    internal PresentationDocument SDKPresentationDocument()
    {
        return (PresentationDocument)this.sdkSlidePart.OpenXmlPackage;
    }

    internal IndentFont? PresentationFontOrThemeFontOrNull(int paraLevel)
    {
        var pDefaultTextStyle = this.SDKPresentationDocument().PresentationPart!.Presentation.DefaultTextStyle;
        if (pDefaultTextStyle != null)
        {
            var pDefaultTextStyleFont = new IndentFonts(pDefaultTextStyle).FontOrNull(paraLevel);
            if (pDefaultTextStyleFont != null)
            {
                return pDefaultTextStyleFont;
            }
        }

        var aTextDefault = this.SDKPresentationDocument().PresentationPart!.ThemePart?.Theme.ObjectDefaults!
            .TextDefault;
        return aTextDefault != null
            ? new IndentFonts(aTextDefault).FontOrNull(paraLevel)
            : null;
    }

    internal string ThemeColorHex(A.SchemeColorValues themeColor)
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

    /// <summary>
    ///     Color hexadecimal representation from Referenced Layout or Master Shape for specified Slide Shape.
    /// </summary>
    internal string? ReferencedShapeColorOrNull(P.Shape slidePShape, int indentLevel)
    {
        var referencedLayoutPShape = this.ReferencedLayoutPShapeOf(slidePShape);
        var indentFont = new IndentFonts(referencedLayoutPShape!.TextBody!.ListStyle!).FontOrNull(indentLevel);
        if (this.HexFromName(indentFont, out var layoutShapeColor))
        {
            return layoutShapeColor;
        }

        var referencedMasterPShape = this.ReferencedMasterPShapeOf(referencedLayoutPShape);
        indentFont = new IndentFonts(referencedMasterPShape!.TextBody!.ListStyle!).FontOrNull(indentLevel);
        if (this.HexFromName(indentFont, out var masterShapeColor))
        {
            return masterShapeColor;
        }

        return null;
    }

    private bool HexFromName(IndentFont? indentFont, out string? referencedShapeColorOrNull)
    {
        if (indentFont == null)
        {
            {
                referencedShapeColorOrNull = null;
                return true;
            }
        }

        if (indentFont.Value.ARgbColorModelHex != null)
        {
            {
                referencedShapeColorOrNull = indentFont.Value.ARgbColorModelHex.Val!.Value;
                return true;
            }
        }

        if (indentFont.Value.ASchemeColor != null)
        {
            {
                referencedShapeColorOrNull = this.ThemeColorHex(indentFont.Value.ASchemeColor.Val!.Value);
                return true;
            }
        }

        if (indentFont.Value.ASystemColor != null)
        {
            {
                referencedShapeColorOrNull = indentFont.Value.ASystemColor.LastColor!;
                return true;
            }
        }

        if (indentFont.Value.APresetColor != null)
        {
            var coloName = indentFont.Value.APresetColor.Val!.Value.ToString();
            {
                referencedShapeColorOrNull = ColorTranslator.HexFromName(coloName);
                return true;
            }
        }

        referencedShapeColorOrNull = null;
        return false;
    }

    /// <summary>
    ///     Tries to get referenced Placeholder Shape located on Slide Layout. Returns <c>NULL</c> if such shape is not found.
    /// </summary>
    private P.Shape? ReferencedLayoutPShapeOf(P.Shape slidePShape)
    {
        var slidePh = slidePShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (slidePh == null)
        {
            return null;
        }

        var layoutPShapes =
            this.sdkSlidePart.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!.Elements<P.Shape>();

        if (ReferencedPShape(layoutPShapes, slidePh, out var shape))
        {
            return shape;
        }

        return null;
    }

    private P.Shape? ReferencedMasterPShapeOf(P.Shape layoutPShape)
    {
        var layoutPh = layoutPShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (layoutPh == null)
        {
            return null;
        }

        var masterPShapes =
            this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.CommonSlideData!.ShapeTree!
                .Elements<P.Shape>();

        if (ReferencedPShape(masterPShapes, layoutPh, out var referencedPShape))
        {
            return referencedPShape;
        }

        // https://answers.microsoft.com/en-us/msoffice/forum/all/placeholder-master/0d51dcec-f982-4098-b6b6-94785304607a?page=3
        if (layoutPh.Index?.Value == 4294967295)
        {
            return masterPShapes.FirstOrDefault(x => x.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                .GetFirstChild<P.PlaceholderShape>()?.Index?.Value == 1);
        }

        return null;
    }

    private static bool ReferencedPShape(IEnumerable<P.Shape> layoutPShapes, P.PlaceholderShape slidePh,
        out P.Shape? shape)
    {
        foreach (var layoutPShape in layoutPShapes)
        {
            var layoutPh = layoutPShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                .GetFirstChild<P.PlaceholderShape>();
            if (layoutPh == null)
            {
                continue;
            }

            if (slidePh.Index is not null && layoutPh.Index is not null &&
                slidePh.Index == layoutPh.Index)
            {
                {
                    shape = layoutPShape;
                    return true;
                }
            }

            if (slidePh.Type == null || layoutPh.Type == null)
            {
                {
                    shape = layoutPShape;
                    return true;
                }
            }

            if (slidePh.Type == P.PlaceholderValues.Body &&
                slidePh.Index is not null && layoutPh.Index is not null)
            {
                if (slidePh.Index == layoutPh.Index)
                {
                    {
                        shape = layoutPShape;
                        return true;
                    }
                }
            }

            var slidePhType = slidePh.Type;
            if (slidePh.Type == P.PlaceholderValues.CenteredTitle)
            {
                slidePhType = P.PlaceholderValues.Title;
            }

            var layoutPhType = layoutPh.Type;
            if (layoutPh.Type == P.PlaceholderValues.CenteredTitle)
            {
                layoutPhType = P.PlaceholderValues.Title;
            }

            if (slidePhType.Equals(layoutPhType))
            {
                {
                    shape = layoutPShape;
                    return true;
                }
            }
        }

        shape = null;
        return false;
    }

    #endregion APIs

    private string GetThemeMappedColor(A.SchemeColorValues themeColor)
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
        var aColorScheme = this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
            .ColorScheme!;
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