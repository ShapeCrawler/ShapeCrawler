using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Groups;

internal sealed class GroupedShapeCollection(IEnumerable<OpenXmlCompositeElement> pGroupElements) : IShapeCollection
{
    public int Count => this.GetGroupedShapes().Count;

    public IShape this[int index] => this.GetGroupedShapes()[index];

    public IShape GetById(int id) => this.GetById<IShape>(id);

    public T GetById<T>(int id)
        where T : IShape => (T)this.GetGroupedShapes().First(shape => shape.Id == id);

    public IShape Shape(string name) => this.GetGroupedShapes().First(shape => shape.Name == name);

    public T Shape<T>(string name)
        where T : IShape => (T)this.GetGroupedShapes().First(shape => shape.Name == name);

    public T Last<T>()
        where T : IShape => (T)this.GetGroupedShapes().Last(shape => shape is T);

    public IEnumerator<IShape> GetEnumerator() => this.GetGroupedShapes().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private List<IShape> GetGroupedShapes()
    {
        var groupedShapes = new List<IShape>();
        foreach (var pGroupShapeElement in pGroupElements)
        {
            IShape? shape = null;
            switch (pGroupShapeElement)
            {
                case P.GroupShape pGroupShape:
                    shape = new Group(
                        new Shape(new Position(pGroupShape), new ShapeSize(pGroupShape), new ShapeId(pGroupShape), pGroupShape),
                        pGroupShape);
                    break;

                case P.Shape { TextBody: not null } pShape:
                    shape = new GroupedShape(
                        new TextShape(
                            new Shape(new Position(pShape), new ShapeSize(pShape), new ShapeId(pShape), pShape),
                            new TextBox(pShape.TextBody)
                        ),
                        pShape
                    );
                    break;

                case P.Shape pShape:
                    shape = new GroupedShape(
                        new Shape(new Position(pShape), new ShapeSize(pShape), new ShapeId(pShape), pShape),
                        pShape
                    );
                    break;

                case P.Picture pPicture:
                    {
                        var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
                        var blipEmbed = aBlip?.Embed;
                        if (blipEmbed is not null)
                        {
                            shape = new Picture(
                                new Shape(new Position(pPicture), new ShapeSize(pPicture), new ShapeId(pPicture), pPicture),
                                pPicture,
                                aBlip!
                            );
                        }

                        break;
                    }

                default:
                    break;
            }

            if (shape != null)
            {
                groupedShapes.Add(shape);
            }
        }

        return groupedShapes;
    }
}