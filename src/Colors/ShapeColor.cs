using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Colors;

internal readonly ref struct ShapeColor
{
    private readonly P.Shape pShape;
    private readonly PresentationColor presColor;

    internal ShapeColor(OpenXmlPart openXmlPart, A.Text aText)
        : this(new PresentationColor(openXmlPart), aText.Ancestors<P.Shape>().First())
    {
    }

    internal ShapeColor(PresentationColor presColor, P.Shape pShape)
    {
        this.pShape = pShape;
        this.presColor = presColor;
    }

    internal string? HexOrNull()
    {
        if (this.pShape.ShapeStyle == null)
        {
            return null;
        }

        var sdkAFontReference = this.pShape.ShapeStyle.FontReference!;
        if (sdkAFontReference.RgbColorModelHex != null)
        {
            return sdkAFontReference.RgbColorModelHex.Val;
        }

        if (sdkAFontReference.SchemeColor != null)
        {
            return this.presColor.ThemeColorHex(sdkAFontReference.SchemeColor.Val!);
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
        if (this.pShape.ShapeStyle == null)
        {
            return null;
        }
        
        var aFontReference = this.pShape.ShapeStyle.FontReference!;
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