using ShapeCrawler.Positions;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class TextShape(P.Shape pShape, TextBox textBox) : Shape(new Position(pShape), new ShapeSize(pShape), new ShapeId(pShape), pShape)
{
    public override ITextBox TextBox => textBox;

    public override void SetText(string text) => textBox.SetText(text);
    
    public override void SetFontName(string fontName)
    {
        foreach (var paragraph in this.TextBox.Paragraphs)
        {
            paragraph.SetFontName(fontName);
        }
    }

    public override void SetFontSize(decimal fontSize)
    {
        foreach (var paragraph in this.TextBox.Paragraphs)
        {
            paragraph.SetFontSize((int)fontSize);
        }
    }

    public override void SetFontColor(string colorHex)
    {
        foreach (var paragraph in this.TextBox.Paragraphs)
        {
            paragraph.SetFontColor(colorHex);
        }
    }
}