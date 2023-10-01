using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

internal sealed class SlideGroupedShapes : IReadOnlyShapes
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    private readonly IEnumerable<OpenXmlCompositeElement> pGroupElements;

    internal SlideGroupedShapes(TypedOpenXmlPart sdkTypedOpenXmlPart, IEnumerable<OpenXmlCompositeElement> pGroupElements)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.pGroupElements = pGroupElements;
    }

    public int Count => this.GroupedShapes().Count;
    public T GetById<T>(int shapeId) where T : IShape => (T)this.GroupedShapes().First(shape => shape.Id == shapeId);
    T IReadOnlyShapes.GetByName<T>(string shapeName) => (T)this.GroupedShapes().First(shape => shape.Name == shapeName);
    public IShape GetByName(string shapeName) => this.GroupedShapes().First(shape => shape.Name == shapeName);
    public T GetByName<T>(string shapeName) => (T)this.GroupedShapes().First(shape => shape.Name == shapeName);
    public IEnumerator<IShape> GetEnumerator() => this.GroupedShapes().GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();
    public IShape this[int index] => this.GroupedShapes()[index];

    private List<IShape> GroupedShapes()
    {
        var groupedShapes = new List<IShape>();
        foreach (var pGroupShapeElement in this.pGroupElements)
        {
            IShape? shape = null;
            if (pGroupShapeElement is P.GroupShape pGroupShape)
            {
                shape = new SlideGroupShape(this.sdkTypedOpenXmlPart, pGroupShape);
            }
            else if (pGroupShapeElement is P.Shape pShape)
            {
                if (pShape.TextBody is not null)
                {
                    shape = new AutoShape(this.sdkTypedOpenXmlPart, pShape, new TextFrame(this.sdkTypedOpenXmlPart, pShape.TextBody));
                }
                else
                {
                    shape = new AutoShape(this.sdkTypedOpenXmlPart, pShape);
                }
            }
            else if (pGroupShapeElement is P.Picture pPicture)
            {
                var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
                var blipEmbed = aBlip?.Embed;
                if (blipEmbed is not null)
                {
                    shape = new SlidePicture(this.sdkTypedOpenXmlPart, pPicture, aBlip!);
                }
            }

            if (shape != null)
            {
                groupedShapes.Add(shape);
            }
        }

        return groupedShapes;
    }
}