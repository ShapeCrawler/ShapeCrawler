using System.Collections.Generic;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a collection of shapes.
/// </summary>
public interface IShapes : IReadOnlyList<IShape>
{
    /// <summary>
    ///     Gets shape by identifier.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T GetById<T>(int id)
        where T : IShape;

    /// <summary>
    ///     Tries to get shape by identifier, returns null if shape is not found.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T? TryGetById<T>(int id) 
        where T : IShape;

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T GetByName<T>(string name) 
        where T : IShape;
    
    /// <summary>
    ///     Tries to get shape by name, returns null if shape is not found.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T? TryGetByName<T>(string name) 
        where T : IShape;

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    IShape GetByName(string name);
    
    /// <summary>
    ///     Gets shape by specified type.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T Last<T>()
        where T : IShape;
}