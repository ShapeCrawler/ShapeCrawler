using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using SkiaSharp;

namespace ShapeCrawler.Media;

/// <summary>
///     Represents a shape containing video content.
/// </summary>
public interface IVideoShape : IShape
{
    /// <summary>
    ///     Gets bytes of video content.
    /// </summary>
    public byte[] BinaryData { get; }

    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    string MIME { get; }
}

internal sealed class SCVideoShape : SCMediaShape, IVideoShape
{
    internal SCVideoShape(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
        : base(pShapeTreeChild, parentSlideObject, parentShapeCollection)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.Video;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}