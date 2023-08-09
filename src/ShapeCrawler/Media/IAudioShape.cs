using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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

internal sealed class SCSlideAudio : SCMediaShape, IAudioShape
{
    internal SCSlideAudio(
        OpenXmlCompositeElement pShapeTree,
        SCSlide slide,
        SCSlideShapes slideShapes,
        SlidePart sdkSlidePart)
        : base(pShapeTree, slide, slideShapes, sdkSlidePart)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.Audio;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }

    internal override IHtmlElement ToHtmlElement()
    {
        throw new System.NotImplementedException();
    }

    internal override string ToJson()
    {
        throw new System.NotImplementedException();
    }
}