using System.IO;

namespace ShapeCrawler;

public sealed partial class Presentation
{
    /// <summary>
    ///     Represents a draft picture.
    /// </summary>
    public sealed class DraftPicture
    {
        internal string DraftName { get; private set; } = "Picture";

        internal int DraftX { get; private set; }

        internal int DraftY { get; private set; }

        internal int DraftWidth { get; private set; } = 100;

        internal int DraftHeight { get; private set; } = 100;

        internal Stream ImageStream { get; private set; } = new MemoryStream();

        internal string? GeometryName { get; private set; }

        /// <summary>
        ///     Sets name.
        /// </summary>
        public DraftPicture Name(string name)
        {
            this.DraftName = name;
            return this;
        }

        /// <summary>
        ///     Sets X-position.
        /// </summary>
        public DraftPicture X(int x)
        {
            this.DraftX = x;
            return this;
        }

        /// <summary>
        ///     Sets Y-position.
        /// </summary>
        public DraftPicture Y(int y)
        {
            this.DraftY = y;
            return this;
        }

        /// <summary>
        ///     Sets width.
        /// </summary>
        public DraftPicture Width(int width)
        {
            this.DraftWidth = width;
            return this;
        }

        /// <summary>
        ///     Sets height.
        /// </summary>
        public DraftPicture Height(int height)
        {
            this.DraftHeight = height;
            return this;
        }

        /// <summary>
        ///     Sets image.
        /// </summary>
        public DraftPicture Image(Stream image)
        {
            this.ImageStream = image;
            return this;
        }

        /// <summary>
        ///     Sets geometry form.
        /// </summary>
        public DraftPicture GeometryType(string geometry)
        {
            this.GeometryName = geometry;
            return this;
        }
    }
}