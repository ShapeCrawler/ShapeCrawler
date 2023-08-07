using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Charts;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a Line Chart.
/// </summary>
public interface ILineChart : IChart
{
}

internal sealed class SCLineChart : SCChart, ILineChart
{
    internal SCLineChart(
        P.GraphicFrame pGraphicFrame, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<ShapeCollection, SCGroupShape> shapeCollectionOf,
        TypedOpenXmlPart slideTypedOpenXmlPart)
        : base(pGraphicFrame, slideOf, shapeCollectionOf, slideTypedOpenXmlPart)
    {
    }
}