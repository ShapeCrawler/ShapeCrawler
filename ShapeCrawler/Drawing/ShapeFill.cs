using System.Drawing;
using System.Globalization;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing
{
    /// <summary>
    ///     Represents a shape fill.
    /// </summary>
    public class ShapeFill : IShapeFill
    {
        /// <summary>
        ///     Returns fill type.
        /// </summary>
        public FillType Type { get; }

        /// <summary>
        ///     Returns picture image. Returns <c>null</c> if fill type is not picture.
        /// </summary>
        public ImageSc Picture { get; }

        /// <summary>
        ///     Returns instance of the <see cref="System.Drawing.Color" />. Returns <c>null</c> if fill type is not solid color.
        /// </summary>
        public Color SolidColor { get; }

        #region Public Methods

        public static ShapeFill FromXmlSolidFill(A.RgbColorModelHex rgbColorModelHex)
        {
            var hexColor = rgbColorModelHex.Val.ToString();
            var hexColorInt = int.Parse(hexColor, NumberStyles.HexNumber, CultureInfo.CurrentCulture);
            Color clr = Color.FromArgb(hexColorInt);

            return new ShapeFill(clr);
        }

        #endregion Public Methods

        internal static ShapeFill FromASchemeClr(A.SchemeColor schemeColor)
        {
            return new ShapeFill(schemeColor);
        }

        #region Constructors

        public ShapeFill(ImageSc image)
        {
            Check.NotNull(image, nameof(image));

            Picture = image;
            Type = FillType.Picture;
        }

        private ShapeFill(Color clr)
        {
            Check.NotNull(clr, nameof(clr));

            SolidColor = clr;
            Type = FillType.Solid;
        }

        private ShapeFill(A.SchemeColor schemeColor)
        {
        }

        #endregion Constructors
    }
}