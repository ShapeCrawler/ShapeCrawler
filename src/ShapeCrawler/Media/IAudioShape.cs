using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Media;
using ShapeCrawler.Shapes;
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

internal sealed class SCAudioShape : SCMediaShape, IAudioShape
{
    internal SCAudioShape(
        OpenXmlCompositeElement pShapeTreesChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
            : base(pShapeTreesChild, parentSlideObject, parentShapeCollection)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.Audio;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}