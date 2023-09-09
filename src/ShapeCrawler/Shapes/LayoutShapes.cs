using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Shapes;

internal sealed record LayoutShapes : IReadOnlyShapes
{
    private readonly SlideLayoutPart sdkSlideLayoutPart;

    internal LayoutShapes(SlideLayoutPart sdkSlideLayoutPart)
    {
        this.sdkSlideLayoutPart = sdkSlideLayoutPart;
    }

    public IEnumerator<IShape> GetEnumerator()
    {
        throw new System.NotImplementedException();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public int Count => this.ShapeList().Count;

    private List<IShape> ShapeList()
    {
        throw new System.NotImplementedException();
    }

    public T GetById<T>(int shapeId) where T : IShape
    {
        throw new System.NotImplementedException();
    }

    public T GetByName<T>(string shapeName) where T : IShape
    {
        throw new System.NotImplementedException();
    }

    public IShape GetByName(string shapeName)
    {
        throw new System.NotImplementedException();
    }

    public IShape this[int index] => this.ShapeList()[index];
}