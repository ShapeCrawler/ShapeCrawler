using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Wrappers;

internal sealed class SdkSlidePShape
{
    private readonly P.Shape sdkPShape;
    private readonly SdkSlidePart sdkSlidePartWrap;

    internal SdkSlidePShape(SdkSlidePart sdkSlidePartWrap, P.Shape sdkPShape)
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

        if (sdkAFontReference.PresetColor != null)
        {
            var coloName = sdkAFontReference.PresetColor.Val!.Value.ToString();
            return ColorTranslator.HexFromName(coloName);
        }

        return null;
    }
}