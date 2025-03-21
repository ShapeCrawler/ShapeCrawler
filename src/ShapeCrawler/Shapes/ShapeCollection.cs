using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.GroupShapes;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class ShapeCollection(OpenXmlPart openXmlPart) : IShapeCollection
{
    public int Count => this.GetShapes().Count();

    public IShape this[int index] => this.GetShapes().ElementAt(index);

    public T GetById<T>(int id)
        where T : IShape => (T)this.GetShapes().First(shape => shape.Id == id);

    public T? TryGetById<T>(int id)
        where T : IShape => (T?)this.GetShapes().FirstOrDefault(shape => shape.Id == id);

    public T GetByName<T>(string name)
        where T : IShape => (T)this.GetByName(name);

    public T? TryGetByName<T>(string name)
        where T : IShape => (T?)this.GetShapes().FirstOrDefault(shape => shape.Name == name);

    public IShape GetByName(string name) =>
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

    private IEnumerable<IShape> GetShapes()
    {
        var pShapeTree = openXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.Slide.CommonSlideData!.ShapeTree!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideLayout.CommonSlideData!.ShapeTree!,
            NotesSlidePart sdkNotesSlidePart => sdkNotesSlidePart.NotesSlide.CommonSlideData!.ShapeTree!,
            _ => ((SlideMasterPart)openXmlPart).SlideMaster.CommonSlideData!.ShapeTree!
        };
        foreach (var pShapeTreeElement in pShapeTree.OfType<OpenXmlCompositeElement>())
        {
            if (pShapeTreeElement is P.GroupShape pGroupShape)
            {
                yield return new GroupShape(pGroupShape);
            }
            else if (pShapeTreeElement is P.ConnectionShape pConnectionShape)
            {
             yield return new SlideLine(pConnectionShape);
            }
            else if (pShapeTreeElement is P.Shape pShape)
            {
                if (pShape.TextBody is not null)
                {
                    yield return 
                        new RootShape(
                            pShape,
                            new AutoShape(
                                pShape,
                                new TextBox(pShape.TextBody)));
                }
                else
                {
                    yield return 
                        new RootShape(
                            pShape,
                            new AutoShape(pShape));
                }
            }
            else if (pShapeTreeElement is P.GraphicFrame pGraphicFrame)
            {
                var aGraphicData = pShapeTreeElement.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>();
                if (aGraphicData!.Uri!.Value!.Equals(
                        "http://schemas.openxmlformats.org/presentationml/2006/ole",
                        StringComparison.Ordinal))
                {
                    yield return new OleObject(pGraphicFrame);
                    continue;
                }

                var pPicture = pShapeTreeElement.Descendants<P.Picture>().FirstOrDefault();
                if (pPicture != null)
                {
                    var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
                    var blipEmbed = aBlip?.Embed;
                    if (blipEmbed is null)
                    {
                        continue;
                    }

                    yield return new Picture(pPicture, aBlip!);
                    continue;
                }

                if (IsChartPGraphicFrame(pShapeTreeElement))
                {
                    aGraphicData = pShapeTreeElement.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !;
                    var cChartRef = aGraphicData.GetFirstChild<C.ChartReference>() !;
                    var sdkChartPart = (ChartPart)openXmlPart.GetPartById(cChartRef.Id!);
                    var cPlotArea = sdkChartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea;
                    var cCharts = cPlotArea!.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
                    pShapeTreeElement.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !
                        .GetFirstChild<C.ChartReference>();
                    pGraphicFrame = (P.GraphicFrame)pShapeTreeElement;
                    if (cCharts.Count() > 1)
                    {
                        // Combination chart
                        yield return new Chart(
                            sdkChartPart,
                            pGraphicFrame,
                            new Categories(sdkChartPart, cCharts));
                        continue;
                    }

                    var chartType = cCharts.Single().LocalName;

                    if (chartType is "lineChart" or "barChart" or "pieChart")
                    {
                        yield return new Chart(
                            sdkChartPart,
                            pGraphicFrame,
                            new Categories(sdkChartPart, cCharts));
                        continue;
                    }

                    if (chartType is "scatterChart" or "bubbleChart")
                    {
                        yield return new Chart(
                            sdkChartPart,
                            pGraphicFrame,
                            new NullCategories());
                        continue;
                    }

                    yield return new Chart(
                        sdkChartPart,
                        pGraphicFrame,
                        new Categories(sdkChartPart, cCharts));
                }
                else if (IsTablePGraphicFrame(pShapeTreeElement))
                {
                    yield return new Table(pShapeTreeElement);
                }
            }
            else if (pShapeTreeElement is P.Picture pPicture)
            {
                var element = pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.ChildElements
                    .FirstOrDefault();

                switch (element)
                {
                    case A.AudioFromFile:
                        {
                            var aAudioFile = pPicture.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties
                                .GetFirstChild<A.AudioFromFile>();
                            if (aAudioFile is not null)
                            {
                                yield return new MediaShape(pPicture);
                            }

                            continue;
                        }

                    case A.VideoFromFile:
                        {
                            yield return new MediaShape(pPicture);
                            continue;
                        }
                }

                var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
                var blipEmbed = aBlip?.Embed;
                if (blipEmbed is null)
                {
                    continue;
                }

                yield return new Picture(pPicture, aBlip!);
            }
        }
    }
}