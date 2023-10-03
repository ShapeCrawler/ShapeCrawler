using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.ShapesCollection;
using ShapeCrawler.SlideShape;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;
using Picture = ShapeCrawler.Drawing.Picture;

namespace ShapeCrawler.Shapes;

internal sealed class GroupedShapeList : IShapeList
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    private readonly IEnumerable<OpenXmlCompositeElement> pGroupElements;

    internal GroupedShapeList(TypedOpenXmlPart sdkTypedOpenXmlPart,
        IEnumerable<OpenXmlCompositeElement> pGroupElements)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.pGroupElements = pGroupElements;
    }

    public int Count => this.GroupedShapes().Count;
    public T GetById<T>(int id) where T : IShape => (T)this.GroupedShapes().First(shape => shape.Id == id);
    T IShapeList.GetByName<T>(string name) => (T)this.GroupedShapes().First(shape => shape.Name == name);
    public IShape GetByName(string name) => this.GroupedShapes().First(shape => shape.Name == name);
    public T GetByName<T>(string name) => (T)this.GroupedShapes().First(shape => shape.Name == name);
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
                shape = new GroupShape(this.sdkTypedOpenXmlPart, pGroupShape);
            }
            else if (pGroupShapeElement is P.Shape pShape)
            {
                if (pShape.TextBody is not null)
                {
                    shape = new GroupedShape(
                        this.sdkTypedOpenXmlPart,
                        pShape,
                        new AutoShape(this.sdkTypedOpenXmlPart, pShape,
                            new TextFrame(this.sdkTypedOpenXmlPart, pShape.TextBody))
                    );
                }
                else
                {
                    shape = new GroupedShape(
                        this.sdkTypedOpenXmlPart,
                        pShape,
                        new AutoShape(this.sdkTypedOpenXmlPart, pShape)
                    );
                }
            }
            else if (pGroupShapeElement is P.Picture pPicture)
            {
                var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
                var blipEmbed = aBlip?.Embed;
                if (blipEmbed is not null)
                {
                    shape = new Picture(this.sdkTypedOpenXmlPart, pPicture, aBlip!);
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