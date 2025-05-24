using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Positions;
using ShapeCrawler.Slides;
using ShapeCrawler.Tables;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Shapes;

internal sealed class ShapeCollection(OpenXmlPart openXmlPart) : IShapeCollection
{
    public int Count => this.GetShapes().Count();

    public IShape this[int index] => this.GetShapes().ElementAt(index);

    public IShape GetById(int id) => this.GetById<IShape>(id);

    public T GetById<T>(int id)
        where T : IShape => (T)this.GetShapes().First(shape => shape.Id == id);

    public T? TryGetById<T>(int id)
        where T : IShape => (T?)this.GetShapes().FirstOrDefault(shape => shape.Id == id);

    public T GetByName<T>(string name)
        where T : IShape => (T)this.Shape(name);

    public T Shape<T>(string name)
        where T : IShape => (T)this.GetShapes().First(shape => shape.Name == name);

    public IShape Shape(string name) =>
        this.GetShapes().FirstOrDefault(shape => shape.Name == name)
        ?? throw new SCException("Shape not found");

    public T Last<T>()
        where T : IShape => (T)this.GetShapes().Last(shape => shape is T);

    public IEnumerator<IShape> GetEnumerator() => this.GetShapes().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private static bool IsTablePGraphicFrame(OpenXmlCompositeElement pShapeTreeChild)
    {
        if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
        {
            var graphicData = pGraphicFrame.Graphic!.GraphicData!;
            if (graphicData.Uri!.Value!.Equals(
                    "http://schemas.openxmlformats.org/drawingml/2006/table",
                    StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsChartPGraphicFrame(OpenXmlCompositeElement pShapeTreeChild)
    {
        if (pShapeTreeChild is P.GraphicFrame)
        {
            var aGraphicData = pShapeTreeChild.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !;
            if (aGraphicData.Uri!.Value!.Equals(
                    "http://schemas.openxmlformats.org/drawingml/2006/chart",
                    StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<IShape> CreateConnectionShape(P.ConnectionShape pConnectionShape)
    {
        yield return new SlideLine(
            new Shape(new Position(pConnectionShape), new ShapeSize(pConnectionShape), new ShapeId(pConnectionShape), pConnectionShape),
            pConnectionShape
        );
    }

    private static IEnumerable<IShape> CreateGroupShape(P.GroupShape pGroupShape)
    {
        yield return new Group(
            new Shape(new Position(pGroupShape), new ShapeSize(pGroupShape), new ShapeId(pGroupShape), pGroupShape),
            pGroupShape
        );
    }

    private static IEnumerable<IShape> CreateShape(P.Shape pShape)
    {
        if (pShape.TextBody is not null)
        {
            yield return new TextShape(
                new Shape(new Position(pShape), new ShapeSize(pShape), new ShapeId(pShape), pShape),
                new TextBox(pShape.TextBody)
            );
        }
        else
        {
            yield return new Shape(new Position(pShape), new ShapeSize(pShape), new ShapeId(pShape), pShape);
        }
    }

    // ReSharper disable once InconsistentNaming
    private static bool IsOLEObject(A.GraphicData aGraphicData) =>
        aGraphicData.Uri?.Value?.Equals(
            "http://schemas.openxmlformats.org/presentationml/2006/ole",
            StringComparison.Ordinal) ?? false;

    private static IEnumerable<IShape> CreatePictureShapes(P.Picture pPicture)
    {
        var element = pPicture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?
            .ChildElements.FirstOrDefault();

        if (element is A.AudioFromFile or A.VideoFromFile)
        {
            yield return new MediaShape(
                new Shape(new Position(pPicture), new ShapeSize(pPicture), new ShapeId(pPicture), pPicture),
                new SlideShapeOutline(pPicture.ShapeProperties!),
                new ShapeFill(pPicture.ShapeProperties!),
                pPicture
            );
            yield break;
        }
        
        var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
        if (aBlip?.Embed != null)
        {
            yield return new Picture(
                new Shape(new Position(pPicture), new ShapeSize(pPicture), new ShapeId(pPicture), pPicture),
                pPicture,
                aBlip
            );
        }
    }

    private IEnumerable<IShape> GetShapes()
    {
        var pShapeTree = this.GetShapeTreeFromPart();

        foreach (var element in pShapeTree.OfType<OpenXmlCompositeElement>())
        {
            foreach (var shape in this.CreateShapesFromElement(element))
            {
                yield return shape;
            }
        }
    }

    private OpenXmlElement GetShapeTreeFromPart() => openXmlPart switch
    {
        SlidePart sdkSlidePart => sdkSlidePart.Slide.CommonSlideData!.ShapeTree!,
        SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideLayout.CommonSlideData!.ShapeTree!,
        NotesSlidePart sdkNotesSlidePart => sdkNotesSlidePart.NotesSlide.CommonSlideData!.ShapeTree!,
        _ => ((SlideMasterPart)openXmlPart).SlideMaster.CommonSlideData!.ShapeTree!
    };

    private IEnumerable<IShape> CreateShapesFromElement(OpenXmlCompositeElement element)
    {
        return element switch
        {
            P.GroupShape pGroupShape => CreateGroupShape(pGroupShape),
            P.ConnectionShape pConnectionShape => CreateConnectionShape(pConnectionShape),
            P.Shape pShape => CreateShape(pShape),
            P.GraphicFrame pGraphicFrame => this.CreateGraphicFrameShapes(pGraphicFrame),
            P.Picture pPicture => CreatePictureShapes(pPicture),
            _ => []
        };
    }

    private IEnumerable<IShape> CreateGraphicFrameShapes(P.GraphicFrame pGraphicFrame)
    {
        var aGraphicData = pGraphicFrame.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>();
        if (aGraphicData == null)
        {
            yield break;
        }

        if (IsOLEObject(aGraphicData))
        {
            var pShapeProperties = pGraphicFrame.Descendants<P.ShapeProperties>().First();
            yield return new OLEObject(
                new Shape(new Position(pGraphicFrame), new ShapeSize(pGraphicFrame), new ShapeId(pGraphicFrame), pGraphicFrame),
                new SlideShapeOutline(pShapeProperties),
                new ShapeFill(pShapeProperties),
                pGraphicFrame
            );
            yield break;
        }

        // Check for Picture
        var pPicture = pGraphicFrame.Descendants<P.Picture>().FirstOrDefault();
        if (pPicture != null)
        {
            var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
            if (aBlip?.Embed != null)
            {
                yield return new Picture(
                    new Shape(new Position(pPicture), new ShapeSize(pPicture), new ShapeId(pPicture), pPicture),
                    pPicture,
                    aBlip
                );
            }

            yield break;
        }

        if (IsChartPGraphicFrame(pGraphicFrame))
        {
            foreach (var chart in this.CreateChartShapes(pGraphicFrame))
            {
                yield return chart;
            }

            yield break;
        }

        if (IsTablePGraphicFrame(pGraphicFrame))
        {
            var aTable = pGraphicFrame.GetFirstChild<A.Graphic>()!.GetFirstChild<A.GraphicData>()!
                .GetFirstChild<A.Table>() !;
            yield return new Table(
                new Shape(new Position(pGraphicFrame), new ShapeSize(pGraphicFrame), new ShapeId(pGraphicFrame), pGraphicFrame),
                new TableRowCollection(pGraphicFrame),
                new TableColumnCollection(pGraphicFrame),
                new TableStyleOptions(aTable.TableProperties!),
                pGraphicFrame
            );
        }
    }

    [SuppressMessage("ReSharper", "InconsistentNaming")]
    private IEnumerable<IShape> CreateChartShapes(P.GraphicFrame pGraphicFrame)
    {
        var aGraphicData = pGraphicFrame.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !;
        var cChartRef = aGraphicData.GetFirstChild<C.ChartReference>() !;
        var chartPart = (ChartPart)openXmlPart.GetPartById(cChartRef.Id!);
        var cPlotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea;
        var cCharts = cPlotArea!.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        
        if (cCharts.Count() > 1) // combination chart has multiple chart types
        {
            var cShapeProperties = chartPart.ChartSpace.GetFirstChild<C.ShapeProperties>() !;
            var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
            var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
            yield return new AxisChart(
                new CategoryChart(
                    new Chart(
                        new Shape(
                            new Position(pGraphicFrame),
                            new ShapeSize(pGraphicFrame),
                            new ShapeId(pGraphicFrame),
                            pGraphicFrame
                        ),
                        new SeriesCollection(chartPart, cXCharts),
                        new SlideShapeOutline(cShapeProperties),
                        new ShapeFill(cShapeProperties),
                        chartPart
                    ),
                    new Categories(chartPart)
                ),
                new XAxis(chartPart)
            );

            yield break;
        }

        var chartTypeName = cCharts.Single().LocalName;

        // With axis and categories
        if (chartTypeName is "lineChart" or "barChart")
        {
            var cShapeProperties = chartPart.ChartSpace.GetFirstChild<C.ShapeProperties>() !;
            var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
            var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
            yield return new AxisChart(
                new CategoryChart(
                    new Chart(
                        new Shape(
                            new Position(pGraphicFrame),
                            new ShapeSize(pGraphicFrame),
                            new ShapeId(pGraphicFrame),
                            pGraphicFrame
                        ),
                        new SeriesCollection(chartPart, cXCharts),
                        new SlideShapeOutline(cShapeProperties),
                        new ShapeFill(cShapeProperties),
                        chartPart
                    ),
                    new Categories(chartPart)
                ),
                new XAxis(chartPart)
            );

            yield break;
        }

        // With categories
        if (chartTypeName is "pieChart")
        {
            var cShapeProperties = chartPart.ChartSpace.GetFirstChild<C.ShapeProperties>() !;
            var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
            var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
            yield return new CategoryChart(
                new Chart(
                    new Shape(
                        new Position(pGraphicFrame),
                        new ShapeSize(pGraphicFrame),
                        new ShapeId(pGraphicFrame),
                        pGraphicFrame
                    ),
                    new SeriesCollection(chartPart, cXCharts),
                    new SlideShapeOutline(cShapeProperties),
                    new ShapeFill(cShapeProperties),
                    chartPart
                ),
                new Categories(chartPart)
            );

            yield break;
        }

        // With axis
        if (chartTypeName is "scatterChart" or "bubbleChart")
        {
            var cShapeProperties = chartPart.ChartSpace.GetFirstChild<C.ShapeProperties>() !;
            var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
            var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
            yield return new AxisChart(
                new Chart(
                    new Shape(
                        new Position(pGraphicFrame),
                        new ShapeSize(pGraphicFrame),
                        new ShapeId(pGraphicFrame),
                        pGraphicFrame
                    ),
                    new SeriesCollection(chartPart, cXCharts),
                    new SlideShapeOutline(cShapeProperties),
                    new ShapeFill(cShapeProperties),
                    chartPart
                ),
                new XAxis(chartPart)
            );
            yield break;
        }

        // Other
        var otherChartCShapeProperties = chartPart.ChartSpace.GetFirstChild<C.ShapeProperties>() !;
        var otherChartPlotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
        var otherChartCXCharts = otherChartPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        yield return new Chart(
            new Shape(
                new Position(pGraphicFrame),
                new ShapeSize(pGraphicFrame),
                new ShapeId(pGraphicFrame),
                pGraphicFrame
            ),
            new SeriesCollection(chartPart, otherChartCXCharts),
            new SlideShapeOutline(otherChartCShapeProperties),
            new ShapeFill(otherChartCShapeProperties),
            chartPart
        );
    }
}