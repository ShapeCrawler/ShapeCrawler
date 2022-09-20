using System;

namespace ShapeCrawler.Placeholders
{
    /// <summary>
    ///     Represents placeholder data.
    /// </summary>
    internal class PlaceholderData : IEquatable<PlaceholderData>
    {
        #region Properties

        /// <summary>
        ///     Gets or sets placeholder type.
        /// </summary>
        internal PlaceholderType PlaceholderType { get; set; }

        /// <summary>
        ///     Gets or sets index (p:ph idx="12345").
        /// </summary>
        /// <returns>Index value or null if such index not exist.</returns>
        internal int? Index { get; set; }

        #endregion Properties

        #region Public Methods

        public bool Equals(PlaceholderData other)
        {
            if (other == null)
            {
                return false;
            }

            if (this.PlaceholderType != PlaceholderType.Custom && other.PlaceholderType != PlaceholderType.Custom)
            {
                return this.PlaceholderType == other.PlaceholderType;
            }

            if (this.PlaceholderType == PlaceholderType.Custom && other.PlaceholderType == PlaceholderType.Custom)
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

            var ph = (PlaceholderData) obj;

            return this.Equals(ph);
        }

        /// <summary>
        ///     Returns the hash calculating upon the formula suggested here: https://stackoverflow.com/a/263416/2948684
        /// </summary>
        public override int GetHashCode()
        {
            var hash = 17;
            hash = hash * 23 + this.PlaceholderType.GetHashCode();
            if (this.PlaceholderType == PlaceholderType.Custom)
            {
                hash = hash * 23 + this.Index.GetHashCode();
            }

            return hash;
        }

        #endregion Public Methods
    }
}