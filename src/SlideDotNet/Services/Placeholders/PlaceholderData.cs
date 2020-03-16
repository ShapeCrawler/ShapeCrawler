using System;
using System.Diagnostics;
using SlideDotNet.Enums;

namespace SlideDotNet.Services.Placeholders
{
    /// <summary>
    /// Represents placeholder data.
    /// </summary>
    public class PlaceholderData : IEquatable<PlaceholderData>
    {
        #region Properties

        /// <summary>
        /// Gets or sets placeholder type.
        /// </summary>
        public PlaceholderType PlaceholderType { get; set; }

        /// <summary>
        /// Gets or sets index (p:ph idx="12345").  
        /// </summary>
        /// <returns>Index value or null if such index not exist.</returns>
        public int? Index { get; set; }

        #endregion Properties

        #region Public Methods

        public bool Equals(PlaceholderData other)
        {
            if (other == null)
            {
                return false;
            }

            if (PlaceholderType != PlaceholderType.Custom && other.PlaceholderType != PlaceholderType.Custom)
            {
                return PlaceholderType == other.PlaceholderType;
            }

            if (PlaceholderType == PlaceholderType.Custom && other.PlaceholderType == PlaceholderType.Custom)
            {
                return Index == other.Index;
            }

            return false;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }

            var ph = (PlaceholderData)obj;

            return Equals(ph);
        }

        /// <summary>
        /// Returns the hash calculating upon the formula suggested here: https://stackoverflow.com/a/263416/2948684
        /// </summary>
        /// <remarks></remarks>
        public override int GetHashCode()
        {
            var hash = 17;
            hash = hash * 23 + PlaceholderType.GetHashCode(); //TODO; make readonly
            if (PlaceholderType == PlaceholderType.Custom)
            {
                hash = hash * 23 + Index.GetHashCode();
            }

            return hash;
        }

        #endregion Public Methods
    }
}