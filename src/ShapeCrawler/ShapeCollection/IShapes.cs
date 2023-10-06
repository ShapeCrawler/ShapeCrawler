using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents collection of grouped shapes.
/// </summary>
public interface IShapes : IReadOnlyList<IShape>
{
    /// <summary>
    ///     Gets shape by identifier.
    /// </summary>
    T GetById<T>(int id) where T : IShape;

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    T GetByName<T>(string name) where T : IShape;

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    IShape GetByName(string name);
}