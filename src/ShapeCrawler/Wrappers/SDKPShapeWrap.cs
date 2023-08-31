using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Wrappers;

internal sealed class SDKPShapeWrap
{
    private readonly P.Shape sdkPShape;
    private readonly SDKSlidePartWrap sdkSlidePartWrap;

    internal SDKPShapeWrap(SDKSlidePartWrap sdkSlidePartWrap, P.Shape sdkPShape)
    {
        this.sdkPShape = sdkPShape;
        this.sdkSlidePartWrap = sdkSlidePartWrap;
    }

    internal string? FontColorHexOrNull()
    {
        if (this.sdkPShape.ShapeStyle == null)
        {
            return null;
        }

        var sdkAFontReference = sdkPShape.ShapeStyle.FontReference!;
        if (sdkAFontReference.RgbColorModelHex != null)
        {
            return sdkAFontReference.RgbColorModelHex.Val;
        }

        if (sdkAFontReference.SchemeColor != null)
        {
            return this.sdkSlidePartWrap.ThemeColorHex(sdkAFontReference.SchemeColor.Val!);
        }

        // var aFontReference = sdkPShape.ShapeStyle.FontReference!;
        // var fontReferenceFontData = new ParagraphLevelFont
        // {
        //     ARgbColorModelHex = aFontReference.RgbColorModelHex,
        //     ASchemeColor = aFontReference.SchemeColor,
        //     APresetColor = aFontReference.PresetColor
        // };
        // return this.TryFromParaLevelFont(fontReferenceFontData);

        return null;
    }
}