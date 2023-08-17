using System.Collections;
using System.Collections.Generic;
using OneOf;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents collection of grouped shapes.
/// </summary>
public interface IReadOnlyShapeCollection : IReadOnlyCollection<IShape>
{
    /// <summary>
    ///     Gets shape by identifier.
    /// </summary>
    T GetById<T>(int shapeId) where T : IShape;

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    T GetByName<T>(string shapeName) where T : IShape;

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    IShape GetByName(string shapeName);
}

internal abstract class ReadOnlyShapeCollection : IReadOnlyShapeCollection
{
    internal abstract SlideMaster SlideMaster();
    public abstract IEnumerator<IShape> GetEnumerator();

    public abstract int Count { get; }
    public abstract T GetById<T>(int shapeId) where T : IShape;

    public abstract T GetByName<T>(string shapeName) where T : IShape;

    public abstract IShape GetByName(string shapeName);

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
}