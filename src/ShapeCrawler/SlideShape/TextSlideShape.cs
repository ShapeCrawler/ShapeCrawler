using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

internal sealed class TextSlideShape : Shape
{
    internal TextSlideShape(SlidePart sdkSlidePart, P.Shape pShape)
        : base(pShape)
    {
        this.TextFrame = new TextFrame(sdkSlidePart, pShape.TextBody!);
    }

    public override SCShapeType ShapeType { get; }
    public override bool IsTextHolder => true;
    public override ITextFrame TextFrame { get; }
}