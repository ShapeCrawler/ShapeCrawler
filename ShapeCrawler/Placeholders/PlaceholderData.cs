using System;

namespace ShapeCrawler.Placeholders;

internal class PlaceholderData : IEquatable<PlaceholderData>
{
    #region Properties

    /// <summary>
    ///     Gets or sets placeholder type.
    /// </summary>
    internal SCPlaceholderType PlaceholderType { get; set; }

    /// <summary>
    ///     Gets or sets index (p:ph idx="12345").
    /// </summary>
    /// <returns>Index value or null if such index not exist.</returns>
    internal int? Index { get; set; }

    #endregion Properties

    #region Public Methods

    public bool Equals(PlaceholderData? other)
    {
        if (other == null)
        {
            return false;
        }

        if (this.PlaceholderType != SCPlaceholderType.Custom && other.PlaceholderType != SCPlaceholderType.Custom)
        {
            return this.PlaceholderType == other.PlaceholderType;
        }

        if (this.PlaceholderType == SCPlaceholderType.Custom && other.PlaceholderType == SCPlaceholderType.Custom)
        {
            return this.Index == other.Index;
        }

        return false;
    }

    public override bool Equals(object? obj)
    {
        if (obj == null)
        {
            return false;
        }

        var ph = (PlaceholderData)obj;

        return this.Equals(ph);
    }

    public override int GetHashCode()
    {
        var hash = 17;
        hash = hash * 23 + this.PlaceholderType.GetHashCode();
        if (this.PlaceholderType == SCPlaceholderType.Custom)
        {
            hash = hash * 23 + this.Index.GetHashCode();
        }

        return hash;
    }

    #endregion Public Methods
}