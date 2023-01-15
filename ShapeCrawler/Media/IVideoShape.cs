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

internal sealed class VideoSCShape : SCMediaSCShape, IVideoShape
{
    internal VideoSCShape(OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide, OpenXmlCompositeElement pShapeTreeChild)
        : base(pShapeTreeChild, oneOfSlide, null)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.VideoShape;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}