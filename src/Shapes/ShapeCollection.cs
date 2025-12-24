using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Groups;
using ShapeCrawler.MediaContent;
using ShapeCrawler.Positions;
using ShapeCrawler.Slides;
using ShapeCrawler.SmartArts;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class ShapeCollection(OpenXmlPart openXmlPart) : IShapeCollection
{
    public int Count => this.GetInternalShapes().Count();

    public IShape this[int index] => this.GetInternalShapes().ElementAt(index);

    public IShape GetById(int id) => this.GetById<IShape>(id);

    public T GetById<T>(int id)
        where T : IShape => (T)this.GetShapes().First(shape => shape.Id == id);

    public T Shape<T>(string name)
        where T : IShape
    {
        var shape = this.GetShapes().FirstOrDefault(shape => shape.Name == name) ??
                    throw new SCException("Shape not found");
        return (T)shape;
    }

    public IShape Shape(string name) =>
        this.GetShapes().FirstOrDefault(shape => shape.Name == name)
        ?? throw new SCException("Shape not found");

    public T Last<T>()
        where T : IShape => (T)this.GetShapes().Last(shape => shape is T);

    public IEnumerator<IShape> GetEnumerator() => this.GetInternalShapes().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    internal IEnumerable<DrawingShape> GetInternalShapes()
    {
        var pShapeTree = openXmlPart switch
        {
            SlidePart slidePart => slidePart.Slide.CommonSlideData!.ShapeTree!,
            SlideLayoutPart slideLayoutPart => slideLayoutPart.SlideLayout.CommonSlideData!.ShapeTree!,
            NotesSlidePart notesSlidePart => notesSlidePart.NotesSlide.CommonSlideData!.ShapeTree!,
            _ => ((SlideMasterPart)openXmlPart).SlideMaster.CommonSlideData!.ShapeTree!
        };

        foreach (var element in pShapeTree.OfType<OpenXmlCompositeElement>())
        {
            foreach (var shape in this.CreateShapesFromElement(element))
            {
                yield return shape;
            }
        }
    }

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
            var aGraphicData = pShapeTreeChild.GetFirstChild<A.Graphic>()!.GetFirstChild<A.GraphicData>()!;
            if (aGraphicData.Uri!.Value!.Equals(
                    "http://schemas.openxmlformats.org/drawingml/2006/chart",
                    StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsSmartArtPGraphicFrame(OpenXmlCompositeElement pShapeTreeChild)
    {
        if (pShapeTreeChild is P.GraphicFrame)
        {
            var aGraphicData = pShapeTreeChild.GetFirstChild<A.Graphic>()!.GetFirstChild<A.GraphicData>()!;
            if (aGraphicData.Uri!.Value!.Equals(
                    "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                    StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<DrawingShape> CreateLineShapes(P.ConnectionShape pConnectionShape)
    {
        yield return new LineShape(
            new Position(pConnectionShape),
            new ShapeSize(pConnectionShape),
            new ShapeId(pConnectionShape),
            pConnectionShape
        );
    }

    private static IEnumerable<DrawingShape> CreateGroupShapes(P.GroupShape pGroupShape)
    {
        yield return new GroupShape(pGroupShape);
    }

    private static IEnumerable<DrawingShape> CreateShapes(P.Shape pShape)
    {
        if (pShape.TextBody is not null)
        {
            yield return new TextShape(
                pShape,
                new DrawingTextBox(new TextBoxMargins(pShape.TextBody), pShape.TextBody)
            );
        }
        else
        {
            yield return new DrawingShape(new Position(pShape), new ShapeSize(pShape), new ShapeId(pShape), pShape);
        }
    }

    private static bool IsOleObject(A.GraphicData aGraphicData) =>
        aGraphicData.Uri?.Value?.Equals(
            "http://schemas.openxmlformats.org/presentationml/2006/ole",
            StringComparison.Ordinal) ?? false;

    private static IEnumerable<DrawingShape> CreatePictureShapes(P.Picture pPicture)
    {
        var element = pPicture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?
            .ChildElements.FirstOrDefault();

        if (element is A.AudioFromFile or A.VideoFromFile)
        {
            yield return new MediaShape(
                new Position(pPicture),
                new ShapeSize(pPicture),
                new ShapeId(pPicture),
                pPicture
            );
            yield break;
        }

        var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
        if (aBlip?.Embed != null)
        {
            yield return new PictureShape(
                new Picture(pPicture, aBlip),
                pPicture
            );
        }
    }

    private IEnumerable<IShape> GetShapes() => GetInternalShapes();

    private IEnumerable<DrawingShape> CreateShapesFromElement(OpenXmlCompositeElement element)
    {
        return element switch
        {
            P.GroupShape pGroupShape => CreateGroupShapes(pGroupShape),
            P.ConnectionShape pConnectionShape => CreateLineShapes(pConnectionShape),
            P.Shape pShape => CreateShapes(pShape),
            P.GraphicFrame pGraphicFrame => this.CreateGraphicFrameShapes(pGraphicFrame),
            P.Picture pPicture => CreatePictureShapes(pPicture),
            _ => []
        };
    }

    private IEnumerable<DrawingShape> CreateGraphicFrameShapes(P.GraphicFrame pGraphicFrame)
    {
        var aGraphicData = pGraphicFrame.GetFirstChild<A.Graphic>()!.GetFirstChild<A.GraphicData>();
        if (aGraphicData == null)
        {
            yield break;
        }

        if (IsOleObject(aGraphicData))
        {
            yield return new OleObjectShape(
                new Position(pGraphicFrame),
                new ShapeSize(pGraphicFrame),
                new ShapeId(pGraphicFrame),
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
                yield return new PictureShape(new Picture(pPicture, aBlip), pPicture);
            }

            yield break;
        }

        if (IsChartPGraphicFrame(pGraphicFrame))
        {
            yield return this.CreateChart(pGraphicFrame);
            yield break;
        }

        if (IsSmartArtPGraphicFrame(pGraphicFrame))
        {
            yield return new SmartArtShape(
                new Position(pGraphicFrame),
                new ShapeSize(pGraphicFrame),
                new ShapeId(pGraphicFrame),
                pGraphicFrame
            );

            yield break;
        }

        if (IsTablePGraphicFrame(pGraphicFrame))
        {
            yield return new TableShape(
                new Position(pGraphicFrame),
                new ShapeSize(pGraphicFrame),
                new ShapeId(pGraphicFrame),
                pGraphicFrame
            );
        }
    }

    private DrawingShape CreateChart(P.GraphicFrame pGraphicFrame)
    {
        var aGraphicData = pGraphicFrame.GetFirstChild<A.Graphic>()!.GetFirstChild<A.GraphicData>()!;
        var cChartRef = aGraphicData.GetFirstChild<C.ChartReference>()!;
        var chartPart = (ChartPart)openXmlPart.GetPartById(cChartRef.Id!);
        var cPlotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea;
        var cCharts = cPlotArea!.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

        // Combination chart has multiple chart types
        if (cCharts.Count() > 1)
        {
            var cShapeProperties = chartPart.ChartSpace.GetFirstChild<C.ShapeProperties>()!;
            var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!;
            var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));


            return new ChartShape(
                new Chart(
                    new SeriesCollection(chartPart, cXCharts),
                    new SlideShapeOutline(cShapeProperties),
                    new ShapeFill(cShapeProperties),
                    chartPart,
                    new Categories(chartPart),
                    new XAxis(chartPart)
                ),
                pGraphicFrame
            );
        }

        var chartTypeName = cCharts.Single().LocalName;

        // With axis and categories
        if (chartTypeName is "lineChart" or "barChart")
        {
            var cShapeProperties = chartPart!.ChartSpace.GetFirstChild<C.ShapeProperties>()!;
            var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!;
            var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
            return new ChartShape(
                new Chart(
                    new SeriesCollection(chartPart, cXCharts),
                    new SlideShapeOutline(cShapeProperties),
                    new ShapeFill(cShapeProperties),
                    chartPart,
                    new Categories(chartPart),
                    new XAxis(chartPart)
                ),
                pGraphicFrame
            );
        }

        // With categories
        if (chartTypeName is "pieChart")
        {
            var cShapeProperties = chartPart!.ChartSpace.GetFirstChild<C.ShapeProperties>()!;
            var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!;
            var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
            return new ChartShape(
                new Chart(
                    new SeriesCollection(chartPart, cXCharts),
                    new SlideShapeOutline(cShapeProperties),
                    new ShapeFill(cShapeProperties),
                    chartPart,
                    new Categories(chartPart)
                ),
                pGraphicFrame
            );
        }

        // With axis
        if (chartTypeName is "scatterChart" or "bubbleChart")
        {
            var cShapeProperties = chartPart!.ChartSpace.GetFirstChild<C.ShapeProperties>()!;
            var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!;
            var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
            return new ChartShape(
                new Chart(
                    new SeriesCollection(chartPart, cXCharts),
                    new SlideShapeOutline(cShapeProperties),
                    new ShapeFill(cShapeProperties),
                    chartPart,
                    new XAxis(chartPart)
                ),
                pGraphicFrame
            );
        }

        // Other
        var otherChartCShapeProperties = chartPart.ChartSpace.GetFirstChild<C.ShapeProperties>()!;
        var otherChartPlotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!;
        var otherChartCxCharts = otherChartPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

        return new ChartShape(
            new Chart(
                new SeriesCollection(chartPart, otherChartCxCharts),
                new SlideShapeOutline(otherChartCShapeProperties),
                new ShapeFill(otherChartCShapeProperties),
                chartPart,
                new XAxis(chartPart)
            ),
            pGraphicFrame
        );
    }

}