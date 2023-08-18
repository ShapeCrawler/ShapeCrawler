using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
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