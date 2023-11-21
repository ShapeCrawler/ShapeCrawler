using System.Collections.Generic;

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
    /// <typeparam name="T">Shape type.</typeparam>
    T GetById<T>(int id) where T : IShape;

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T GetByName<T>(string name) where T : IShape;

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    IShape GetByName(string name);
}