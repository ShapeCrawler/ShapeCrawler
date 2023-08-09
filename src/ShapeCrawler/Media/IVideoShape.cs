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
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<SCSlideShapes, SCSlideGroupShape> shapeCollectionOf,
        TypedOpenXmlPart slideTypedOpenXmlPart)
        : base(pShapeTreeChild, slideOf, shapeCollectionOf, slideTypedOpenXmlPart)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.Video;

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