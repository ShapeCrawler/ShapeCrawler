using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Wrappers;

internal sealed class SdkSlidePShape
{
    private readonly P.Shape sdkPShape;
    private readonly PresentationColor presentationColorWrap;

    internal SdkSlidePShape(PresentationColor presentationColorWrap, P.Shape sdkPShape)
    {
        this.sdkPShape = sdkPShape;
        this.presentationColorWrap = presentationColorWrap;
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
            return this.presentationColorWrap.ThemeColorHex(sdkAFontReference.SchemeColor.Val!);
        }

        if (sdkAFontReference.PresetColor != null)
        {
            var coloName = sdkAFontReference.PresetColor.Val!.Value.ToString();
            return ColorTranslator.HexFromName(coloName);
        }

        return null;
    }
}