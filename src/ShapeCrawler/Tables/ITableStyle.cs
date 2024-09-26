using System.Collections.Generic;

namespace ShapeCrawler.Tables;

/// <summary>
///     Represents the tablestyle of a table.
/// </summary>
public interface ITableStyle
{
    /// <summary>
    ///     Gets the name of the style.
    /// </summary> 
    public string Name { get; }

    /// <summary>
    ///     Gets the GUID of the style.
    /// </summary>
    public string GUID { get; }
}

internal class TableStyle : ITableStyle
{
    public TableStyle(string name, string guid)
    {
        this.Name = name;
        this.GUID = guid;
    }
    
    public string Name { get; }

    public string GUID { get; }

    public override bool Equals(object? obj)
    {
        return obj is ITableStyle style &&
               this.Name == style.Name &&
               this.GUID == style.GUID;
    }

    public override int GetHashCode()
    {
        int hashCode = 1242478914;
        hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(this.Name);
        hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(this.GUID);
        return hashCode;
    }
}
