using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Media;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using SkiaSharp;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents an audio shape.
/// </summary>
public interface IAudioShape : IShape
{
    /// <summary>
    ///     Gets bytes of audio content.
    /// </summary>
    public byte[] BinaryData { get; }

    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    string MIME { get; }
}

internal class AudioShape : MediaShape, IAudioShape
{
    internal AudioShape(OpenXmlCompositeElement pShapeTreesChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide)
        : base(pShapeTreesChild, oneOfSlide, null)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.AudioShape;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}