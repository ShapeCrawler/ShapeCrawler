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

internal sealed class ShapeCollection(OpenXmlPart openXmlPart): IShapeCollection
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

    private static IEnumerable<IShape> CreateConnectionShape(P.ConnectionShape pConnectionShape)
    {
        yield return new SlideLine(pConnectionShape);
    }

    private static IEnumerable<IShape> CreateShape(P.Shape pShape)
    {
        if (pShape.TextBody is not null)
        {
            yield return new Shape(pShape, new TextBox(pShape.TextBody));
        }
        else
        {
            yield return new Shape(pShape);
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
            P.GroupShape pGroupShape => this.CreateGroupShape(pGroupShape),
            P.ConnectionShape pConnectionShape => CreateConnectionShape(pConnectionShape),
            P.Shape pShape => CreateShape(pShape),
            P.GraphicFrame pGraphicFrame => this.CreateGraphicFrameShapes(pGraphicFrame),
            P.Picture pPicture => this.CreatePictureShapes(pPicture),
            _ => []
        };
    }

    private IEnumerable<IShape> CreateGroupShape(P.GroupShape pGroupShape)
    {
        yield return new GroupShape(pGroupShape);
    }

    private IEnumerable<IShape> CreateGraphicFrameShapes(P.GraphicFrame pGraphicFrame)
    {
        var aGraphicData = pGraphicFrame.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>();
        if (aGraphicData == null)
        {
            yield break;
        }

        // Check for OLE Object
        if (this.IsOleObject(aGraphicData))
        {
            yield return new OleObject(pGraphicFrame);
            yield break;
        }

        // Check for Picture
        var pPicture = pGraphicFrame.Descendants<P.Picture>().FirstOrDefault();
        if (pPicture != null)
        {
            var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
            if (aBlip?.Embed != null)
            {
                yield return new Picture(pPicture, aBlip);
            }

            yield break;
        }

        // Check for Chart
        if (IsChartPGraphicFrame(pGraphicFrame))
        {
            foreach (var chart in this.CreateChartShapes(pGraphicFrame))
            {
                yield return chart;
            }

            yield break;
        }

        // Check for Table
        if (IsTablePGraphicFrame(pGraphicFrame))
        {
            yield return new Table(pGraphicFrame);
        }
    }

    private bool IsOleObject(A.GraphicData aGraphicData) => 
        aGraphicData.Uri?.Value?.Equals(
            "http://schemas.openxmlformats.org/presentationml/2006/ole",
            StringComparison.Ordinal) ?? false;

    private IEnumerable<IShape> CreateChartShapes(P.GraphicFrame pGraphicFrame)
    {
        var aGraphicData = pGraphicFrame.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !;
        var cChartRef = aGraphicData.GetFirstChild<C.ChartReference>() !;
        var sdkChartPart = (ChartPart)openXmlPart.GetPartById(cChartRef.Id!);
        var cPlotArea = sdkChartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea;
        var cCharts = cPlotArea!.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

        // Combination chart with multiple chart types
        if (cCharts.Count() > 1)
        {
            yield return new Chart(
                sdkChartPart,
                pGraphicFrame,
                new Categories(sdkChartPart, cCharts));
            yield break;
        }

        var chartType = cCharts.Single().LocalName;
        
        // Charts with categories
        if (chartType is "lineChart" or "barChart" or "pieChart")
        {
            yield return new Chart(
                sdkChartPart,
                pGraphicFrame,
                new Categories(sdkChartPart, cCharts));
            yield break;
        }

        // Charts without categories
        if (chartType is "scatterChart" or "bubbleChart")
        {
            yield return new Chart(
                sdkChartPart,
                pGraphicFrame,
                new NullCategories());
            yield break;
        }

        // Default chart handling
        yield return new Chart(
            sdkChartPart,
            pGraphicFrame,
            new Categories(sdkChartPart, cCharts));
    }

    private IEnumerable<IShape> CreatePictureShapes(P.Picture pPicture)
    {
        var element = pPicture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?
            .ChildElements.FirstOrDefault();

        // Check for media shapes
        if (element is A.AudioFromFile or A.VideoFromFile)
        {
            yield return new MediaShape(pPicture);
            yield break;
        }

        // Regular picture
        var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
        if (aBlip?.Embed != null)
        {
            yield return new Picture(pPicture, aBlip);
        }
    }
}