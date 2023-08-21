using System.Collections;
using System.Collections.Generic;

namespace ShapeCrawler.Shapes;

internal sealed record LayoutShapes : IReadOnlyShapeCollection
{
    private readonly SlideLayout parentLayout;

    internal LayoutShapes(SlideLayout parentLayout)
    {
        this.parentLayout = parentLayout;
    }
    
    internal SlideMaster SlideMaster()
    {
        return this.parentLayout.SlideMaster();
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