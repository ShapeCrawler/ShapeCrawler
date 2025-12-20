using System;
using ShapeCrawler.Drawing;
using ShapeCrawler.Positions;
using ShapeCrawler.Units;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal class TextShape(P.Shape pShape, DrawingTextBox textBox)
    : DrawingShape(new Position(pShape), new ShapeSize(pShape), new ShapeId(pShape), pShape)
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

    internal override void Render(SKCanvas canvas)
    {
        base.Render(canvas);

        canvas.Save();
        ApplyRotation(canvas);
        textBox.Render(canvas, this.X, this.Y, this.Width, this.Height);
        canvas.Restore();
    }

    private void ApplyRotation(SKCanvas canvas)
    {
        const double epsilon = 1e-6;
        if (Math.Abs(this.Rotation) > epsilon)
        {
            var centerX = this.X + (this.Width / 2);
            var centerY = this.Y + (this.Height / 2);
            canvas.RotateDegrees(
                (float)this.Rotation,
                (float)new Points(centerX).AsPixels(),
                (float)new Points(centerY).AsPixels()
            );
        }
    }
}