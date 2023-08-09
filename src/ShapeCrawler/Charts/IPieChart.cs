using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Charts;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents Pie chart interface.
/// </summary>
public interface IPieChart : IChart
{
}

internal sealed class SCPieChart : SCChart, IPieChart
{
    internal SCPieChart(
        P.GraphicFrame pGraphicFrame, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<SCSlideShapes, SCSlideGroupShape> shapeCollectionOf,
        TypedOpenXmlPart slideTypedOpenXmlPart,
        List<ChartWorkbook> chartWorkbooks)
        : base(pGraphicFrame, slideOf, shapeCollectionOf, slideTypedOpenXmlPart, chartWorkbooks)
    {
    }
}