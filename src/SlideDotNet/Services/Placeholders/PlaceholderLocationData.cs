using DocumentFormat.OpenXml.Drawing;
using SlideDotNet.Enums;
using SlideDotNet.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Services.Placeholders
{
    /// <summary>
    /// Represents placeholder data.
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

        public GeometryType Geometry { get; set; } = GeometryType.Rectangle;

        #endregion

        #region Constructors

        /// <summary>
        /// Creates a new <see cref="PlaceholderLocationData"/> instance from <see cref="PlaceholderData"/>.
        /// </summary>
        public PlaceholderLocationData(PlaceholderData phData)
        {
            Check.NotNull(phData, nameof(phData));

            PlaceholderType = phData.PlaceholderType;
            Index = phData.Index;
        }

        #endregion Constructors
    }
}
