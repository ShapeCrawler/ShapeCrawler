using SlideDotNet.Validation;

namespace SlideDotNet.Services.Placeholders
{
    /// <summary>
    /// Represents placeholder location data.
    /// </summary>
    public class PlaceholderLocationData : PlaceholderData
    {
        #region Properties

        /// <summary>
        /// Gets or sets X-coordinate's value.
        /// </summary>
        public long X { get; set; }

        /// <summary>
        /// Gets or sets Y-coordinate's value.
        /// </summary>
        public long Y { get; set; }

        /// <summary>
        /// Gets or sets width value.
        /// </summary>
        public long Width { get; set; }

        /// <summary>
        /// Gets or sets height value.
        /// </summary>
        public long Height { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Creates a new <see cref="PlaceholderLocationData"/> instance from <see cref="PlaceholderData"/>.
        /// </summary>
        public PlaceholderLocationData(PlaceholderData phXml)
        {
            Check.NotNull(phXml, nameof(phXml));
            PlaceholderType = phXml.PlaceholderType;
            Index = phXml.Index;
        }

        #endregion Constructors
    }
}
