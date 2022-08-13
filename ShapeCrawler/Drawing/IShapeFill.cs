using System.Drawing;
using System.IO;

namespace ShapeCrawler.Drawing
{
    /// <summary>
    ///     Represents a shape fill.
    /// </summary>
    public interface IShapeFill
    {
        /// <summary>
        ///     Gets fill type.
        /// </summary>
        public FillType Type { get; }

        /// <summary>
        ///     Gets picture image. Returns <c>null</c> if fill type is not picture.
        /// </summary>
        public SCImage? Picture { get; }

        /// <summary>
        ///     Gets instance of the <see cref="System.Drawing.Color" />. Returns <c>null</c> if fill type is not solid color.
        /// </summary>
        public Color SolidColor { get; }

        void SetPicture(Stream image);
    }
}