using System.Collections;
using System.Collections.Generic;

namespace ShapeCrawler.Shapes;

internal sealed record LayoutShapes : IReadOnlyShapeCollection
{

    internal LayoutShapes()
    {
    }

    public IEnumerator<IShape> GetEnumerator()
    {
        throw new System.NotImplementedException();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public int Count { get; }
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
}