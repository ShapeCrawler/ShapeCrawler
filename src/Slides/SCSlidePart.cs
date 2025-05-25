using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using Position = ShapeCrawler.Positions.Position;
using Shape = ShapeCrawler.Shapes.Shape;

namespace ShapeCrawler.Slides;

internal readonly ref struct SCSlidePart(SlidePart slidePart)
{
    internal ISmartArt AddSmartArt(int x, int y, int width, int height, SmartArtType smartArtType)
    {
        var pGraphicFrame = new GraphicFrame();

        // Add ID and name properties
        var nvGraphicFrameProperties = new NonVisualGraphicFrameProperties();
        var nonVisualDrawingProperties = new NonVisualDrawingProperties
        {
            Id = this.GetNextShapeId(), Name = $"SmartArt {smartArtType}"
        };
        var nonVisualGraphicFrameDrawingProperties = new NonVisualGraphicFrameDrawingProperties();
        var applicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

        nvGraphicFrameProperties.Append(nonVisualDrawingProperties);
        nvGraphicFrameProperties.Append(nonVisualGraphicFrameDrawingProperties);
        nvGraphicFrameProperties.Append(applicationNonVisualDrawingProperties);
        pGraphicFrame.Append(nvGraphicFrameProperties);

        // Add transform properties
        var transform = new Transform();
        transform.Append(new A.Offset { X = new Points(x).AsEmus(), Y = new Points(y).AsEmus() });
        transform.Append(new A.Extents { Cx = new Points(width).AsEmus(), Cy = new Points(height).AsEmus() });
        pGraphicFrame.Append(transform);

        // Create the diagram graphic
        var graphic = new A.Graphic();
        var graphicData = new A.GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram" };

        // Instead of using Diagram class directly, we'll use a simple approach
        // with just a GraphicData container that identifies as a diagram
        // This will create a valid empty SmartArt shell that can be modified later
        graphic.Append(graphicData);
        pGraphicFrame.Append(graphic);

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pGraphicFrame);

        return
            new SmartArt(
                new Shape(
                    new Position(pGraphicFrame),
                    new ShapeSize(pGraphicFrame),
                    new ShapeId(pGraphicFrame),
                    pGraphicFrame),
                new SmartArtNodeCollection()
            );
    }
    
    private uint GetNextShapeId()
    {
        // Get all existing shape IDs from the slide
        var shapeIds = slidePart.Slide.Descendants<NonVisualDrawingProperties>()
            .Select(p => p.Id?.Value ?? 0)
            .ToList();

        // Find the maximum ID and add 1, or start with 1 if no shapes exist
        return shapeIds.Count > 0 ? shapeIds.Max() + 1 : 1;
    }
}