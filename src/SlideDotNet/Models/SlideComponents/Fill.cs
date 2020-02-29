using System.Drawing;
using System.Globalization;
using SlideDotNet.Enums;
using SlideDotNet.Validation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Models.SlideComponents
{

    public class Fill
    {
        public FillType Type { get; }

        public ImageEx Picture { get; }

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

        public static Fill FromXmlSolidFill(A.SolidFill xmlSldFill)
        {
            var hexColor = xmlSldFill.RgbColorModelHex.Val.ToString();
            var hexColorInt = int.Parse(hexColor, NumberStyles.HexNumber);
            Color clr = Color.FromArgb(hexColorInt);

            return new Fill(clr);
        }
    }
}