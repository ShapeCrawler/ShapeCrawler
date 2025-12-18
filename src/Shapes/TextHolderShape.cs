using ShapeCrawler.Positions;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal class TextHolderShape(P.Shape pShape, ShapeText shapeText) : DrawingShape(new Position(pShape), new ShapeSize(pShape), new ShapeId(pShape), pShape)
{
    public override IShapeText ShapeText => shapeText;

    public override void SetText(string text) => shapeText.SetText(text);
    
    public override void SetFontName(string fontName)
    {
        foreach (var paragraph in this.ShapeText.Paragraphs)
        {
            paragraph.SetFontName(fontName);
        }
    }

    public override void SetFontSize(decimal fontSize)
    {
        foreach (var paragraph in this.ShapeText.Paragraphs)
        {
            paragraph.SetFontSize((int)fontSize);
        }
    }

    public override void SetFontColor(string colorHex)
    {
        foreach (var paragraph in this.ShapeText.Paragraphs)
        {
            paragraph.SetFontColor(colorHex);
        }
    }
}