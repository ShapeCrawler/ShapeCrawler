namespace ShapeCrawler.Shared;

/// <summary>
/// It represents type attribute.
/// </summary>
public interface IEnumeration
{
    /// <summary>
    /// Gets the enum name.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Gets the enum value.
    /// </summary>
    string Value { get; }
}
