using ShapeCrawler.Drawing;
using ShapeCrawler.Groups;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class GroupedTextShape(P.Shape pShape, DrawingTextBox textBox, GroupedShape groupedShape)
    : TextShape(pShape, textBox)
{
    public override decimal X
    {
        get => groupedShape.X;
        set => groupedShape.X = value;
    }

    public override decimal Y
    {
        get => groupedShape.Y;
        set => groupedShape.Y = value;
    }

    public override decimal Width
    {
        get => groupedShape.Width;
        set => groupedShape.Width = value;
    }

    public override decimal Height
    {
        get => groupedShape.Height;
        set => groupedShape.Height = value;
    }
}