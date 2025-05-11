using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Colors;

internal readonly ref struct ShapeColor(PresentationColor presColor, P.Shape pShape)
{
    internal string? HexOrNull()
    {
        if (pShape.ShapeStyle == null)
        {
            return null;
        }

        var sdkAFontReference = pShape.ShapeStyle.FontReference!;
        if (sdkAFontReference.RgbColorModelHex != null)
        {
            return sdkAFontReference.RgbColorModelHex.Val;
        }

        if (sdkAFontReference.SchemeColor != null)
        {
            return presColor.ThemeColorHex(sdkAFontReference.SchemeColor.Val!);
        }

        if (sdkAFontReference.PresetColor != null)
        {
            var coloName = sdkAFontReference.PresetColor.Val!.Value.ToString();
            
            return ColorTranslator.HexFromName(coloName);
        }

        return null;
    }

    internal ColorType? TypeOrNull()
    {
        if (pShape.ShapeStyle == null)
        {
            return null;
        }
        
        var aFontReference = pShape.ShapeStyle.FontReference!;
        if (aFontReference.RgbColorModelHex != null)
        {
            return ColorType.RGB;
        }
        
        if (aFontReference.SchemeColor != null)
        {
            return ColorType.Theme;
        }
        
        if (aFontReference.PresetColor != null)
        {
            return ColorType.Preset;
        }
        
        return null;
    }
}