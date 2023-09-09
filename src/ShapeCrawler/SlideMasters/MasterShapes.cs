using System;
using System.Collections;
using System.Collections.Generic;

namespace ShapeCrawler.SlideMasters;

internal sealed record MasterShapes : IReadOnlyShapes
{
    public MasterShapes(ISlideMaster parentSlideMaster)
    {
        throw new NotImplementedException();
    }

    public IEnumerator<IShape> GetEnumerator()
    {
        throw new NotImplementedException();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public int Count { get; }
    public T GetById<T>(int shapeId) where T : IShape
    {
        throw new NotImplementedException();
    }

    public T GetByName<T>(string shapeName) where T : IShape
    {
        throw new NotImplementedException();
    }

    public IShape GetByName(string shapeName)
    {
        throw new NotImplementedException();
    }

    public IShape this[int index] => throw new NotImplementedException();
}