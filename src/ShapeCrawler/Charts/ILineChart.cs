using System.Collections.Generic;
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

internal sealed class SCSlideLineChart : SCSlideChart, ILineChart
{
    internal SCSlideLineChart(
        P.GraphicFrame pGraphicFrame, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<SCSlideShapes, SCSlideGroupShape> shapeCollectionOf,
        TypedOpenXmlPart slideTypedOpenXmlPart,
        List<ChartWorkbook> chartWorkbooks)
        : base(pGraphicFrame, slideOf, shapeCollectionOf, slideTypedOpenXmlPart, chartWorkbooks)
    {
    }
}