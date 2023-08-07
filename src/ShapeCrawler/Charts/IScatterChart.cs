using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Charts;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a Scatter Chart.
/// </summary>
public interface IScatterChart : IChart
{
}

internal sealed class SCScatterChart : SCChart, IScatterChart
{
    internal SCScatterChart(
        P.GraphicFrame pGraphicFrame, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<ShapeCollection, SCGroupShape> shapeCollectionOf,
        TypedOpenXmlPart slideTypedOpenXmlPart)
        : base(pGraphicFrame, slideOf, shapeCollectionOf, slideTypedOpenXmlPart)
    {
    }
}