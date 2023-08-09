using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a Rectangle shape.
/// </summary>
public interface IRectangle : IAutoShape
{
}

internal sealed class SCRectangle : SCAutoShape, IRectangle
{
    internal SCRectangle(
        P.Shape pShape,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<SCSlideShapes, SCSlideGroupShape> shapeCollectionOf,
        TypedOpenXmlPart slideTypedOpenXmlPart)
        : base(pShape, slideOf, shapeCollectionOf, slideTypedOpenXmlPart)
    {
    }

    internal override IHtmlElement ToHtmlElement()
    {
        throw new System.NotImplementedException();
    }
}