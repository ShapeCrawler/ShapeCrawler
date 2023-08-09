using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Charts;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a Combination chart.
/// </summary>
public interface IComboChart : IChart
{
}

internal sealed class SCSlideComboChart : SCSlideChart, IComboChart
{
    internal SCSlideComboChart(
        P.GraphicFrame pGraphicFrame, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slide,
        OneOf<SCSlideShapes, SCSlideGroupShape> shapeCollectionOf,
        TypedOpenXmlPart slideTypedOpenXmlPart,
        List<ChartWorkbook> chartWorkbooks)
        : base(pGraphicFrame, slide, shapeCollectionOf, slideTypedOpenXmlPart, chartWorkbooks)
    {
    }
}