#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a chart category.
/// </summary>
public interface ICategory
{
    /// <summary>
    ///     Gets a value indicating whether the category has a main category.
    /// </summary>
    public bool HasMainCategory { get; }
 
    /// <summary>
    ///     Gets main category.
    /// </summary>
    public ICategory MainCategory { get; }

    /// <summary>
    ///     Gets or sets category name.
    /// </summary>
    string Name { get; set; }
}