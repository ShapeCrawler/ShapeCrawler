using ShapeCrawler.Shared;

namespace ShapeCrawler.Placeholders
{
    /// <summary>
    ///     Represents placeholder data.
    /// </summary>
    internal class PlaceholderLocationData : PlaceholderData
    {
        public PlaceholderLocationData(PlaceholderData phData)
        {
            Check.NotNull(phData, nameof(phData));

            this.PlaceholderType = phData.PlaceholderType;
            this.Index = phData.Index;
        }

        /// <summary>
        ///     Gets or sets X-coordinate's value.
        /// </summary>
        public int X { get; set; }

        /// <summary>
        ///     Gets or sets Y-coordinate's value.
        /// </summary>
        public int Y { get; set; }

        /// <summary>
        ///     Gets or sets width value.
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        ///     Gets or sets height value.
        /// </summary>
        public int Height { get; set; }

        public SCGeometry Geometry { get; set; } = SCGeometry.Rectangle;
    }
}