using System.Drawing;
using System.Globalization;
using SlideDotNet.Enums;
using SlideDotNet.Validation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Models.SlideComponents
{
    /// <summary>
    /// Represents a shape fill.
    /// </summary>
    public class Fill
    {
        /// <summary>
        /// Returns fill type.
        /// </summary>
        public FillType Type { get; }

        /// <summary>
        /// Returns picture image. Returns <c>null</c> if fill type is not picture.
        /// </summary>
        public ImageEx Picture { get; }

        /// <summary>
        /// Returns instance of the <see cref="System.Drawing.Color"/>. Returns <c>null</c> if fill type is not solid color.
        /// </summary>
        public Color SolidColor { get; }

        #region Constructors

        public Fill(ImageEx image)
        {
            Check.NotNull(image, nameof(image));

            Picture = image;
            Type = FillType.Picture;
        }

        private Fill(Color clr)
        {
            Check.NotNull(clr, nameof(clr));

            SolidColor = clr;
            Type = FillType.Solid;
        }

        #endregion Constructors

        #region Public Methods

        public static Fill FromXmlSolidFill(A.RgbColorModelHex rgbColorModelHex)
        {
            var hexColor = rgbColorModelHex.Val.ToString();
            var hexColorInt = int.Parse(hexColor, NumberStyles.HexNumber);
            Color clr = Color.FromArgb(hexColorInt);

            return new Fill(clr);
        }

        #endregion Public Methods
    }
}