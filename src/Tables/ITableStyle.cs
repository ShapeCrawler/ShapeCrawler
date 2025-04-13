using System.Collections.Generic;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a table style of a table.
/// </summary>
public interface ITableStyle
{
    /// <summary>
    ///     Gets the name of the style.
    /// </summary> 
    public string Name { get; }
}

internal class TableStyle(string name): ITableStyle
{
    public string Name { get; } = name;

    public string Guid { get; init; } = string.Empty;

    public override bool Equals(object? obj)
    {
        return obj is TableStyle style &&
               this.Name == style.Name &&
               this.Guid == style.Guid;
    }

    public override int GetHashCode()
    {
        int hashCode = 1242478914;
        hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(this.Name);
        hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(this.Guid);
        return hashCode;
    }
}
