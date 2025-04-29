using System.Collections.Generic;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a collection of shapes.
/// </summary>
public interface IShapeCollection : IReadOnlyList<IShape>
{
    /// <summary>
    ///     Gets shape by identifier.
    /// </summary>
    IShape GetById(int id);
    
    /// <summary>
    ///     Gets shape by identifier.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T GetById<T>(int id)
        where T : IShape;
    
    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    IShape Shape(string name);
    
    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T Shape<T>(string name) 
        where T : IShape;
    
    /// <summary>
    ///     Gets the latest shape.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T Last<T>()
        where T : IShape;
}