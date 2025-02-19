using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;
using Picture = ShapeCrawler.Drawing.Picture;

namespace ShapeCrawler.GroupShapes;

internal sealed record GroupedShapeCollection : IShapeCollection
{
    private readonly OpenXmlPart openXmlPart;
    private readonly IEnumerable<OpenXmlCompositeElement> pGroupElements;

    internal GroupedShapeCollection(OpenXmlPart openXmlPart, IEnumerable<OpenXmlCompositeElement> pGroupElements)
    {
        this.openXmlPart = openXmlPart;
        this.pGroupElements = pGroupElements;
    }

    public int Count => this.GetGroupedShapes().Count;

    public IShape this[int index] => this.GetGroupedShapes()[index];

    public T GetById<T>(int id)
        where T : IShape => (T)this.GetGroupedShapes().First(shape => shape.Id == id);

    public T? TryGetById<T>(int id) 
        where T : IShape => (T?)this.GetGroupedShapes().FirstOrDefault(shape => shape.Id == id);

    T IShapeCollection.GetByName<T>(string name) => (T)this.GetGroupedShapes().First(shape => shape.Name == name);

    T? IShapeCollection.TryGetByName<T>(string name) 
        where T : default => (T?)this.GetGroupedShapes().FirstOrDefault(shape => shape.Name == name);

    public IShape GetByName(string name) => this.GetGroupedShapes().First(shape => shape.Name == name);
    public T Last<T>() 
        where T : IShape => (T)this.GetGroupedShapes().Last(shape => shape is T);

    public T GetByName<T>(string name) => (T)this.GetGroupedShapes().First(shape => shape.Name == name);

    public IEnumerator<IShape> GetEnumerator() => this.GetGroupedShapes().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private List<IShape> GetGroupedShapes()
    {
        var groupedShapes = new List<IShape>();
        foreach (var pGroupShapeElement in this.pGroupElements)
        {
            IShape? shape = null;
            if (pGroupShapeElement is P.GroupShape pGroupShape)
            {
                shape = new GroupShape(this.openXmlPart, pGroupShape);
            }
            else if (pGroupShapeElement is P.Shape pShape)
            {
                if (pShape.TextBody is not null)
                {
                    shape = new GroupedShape(
                        pShape,
                        new AutoShape(
                            this.openXmlPart,
                            pShape,
                            new TextBox(this.openXmlPart, pShape.TextBody)));
                }
                else
                {
                    shape = new GroupedShape(
                        pShape,
                        new AutoShape(this.openXmlPart, pShape));
                }
            }
            else if (pGroupShapeElement is P.Picture pPicture)
            {
                var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
                var blipEmbed = aBlip?.Embed;
                if (blipEmbed is not null)
                {
                    shape = new Picture(this.openXmlPart, pPicture, aBlip!);
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