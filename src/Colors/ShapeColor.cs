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

        var aFontReference = pShape.ShapeStyle.FontReference!;
        if (aFontReference.RgbColorModelHex != null)
        {
            return aFontReference.RgbColorModelHex.Val;
        }

        if (aFontReference.SchemeColor != null)
        {
            return presColor.ThemeColorHex(aFontReference.SchemeColor.Val!);
        }

        if (aFontReference.PresetColor != null)
        {
            var colorName = aFontReference.PresetColor.Val!.Value.ToString();
            
            return ColorTranslator.HexFromName(colorName);
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