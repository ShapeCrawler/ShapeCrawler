using System.Drawing;
using System.Globalization;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing
{
    internal class ShapeFill : IShapeFill
    {
        private ShapeFill()
        {
            this.Type = FillType.NoFill;
        }
        
        private ShapeFill(SCImage image)
        {
            this.Picture = image;
            this.Type = FillType.Picture;
        }
        
        private ShapeFill(Color color)
        {
            this.SolidColor = color;
            this.Type = FillType.Solid;
        }

        private ShapeFill(A.SchemeColor schemeColor)
        {
        }
        
        public FillType Type { get; }

        public SCImage? Picture { get; }

        public Color SolidColor { get; }
        
        public void SetPicture(Stream image)
        {
            if (this.Type == FillType.NoFill)
            {
                // this.
            }
        }

        internal static ShapeFill FromXmlSolidFill(A.RgbColorModelHex rgbColorModelHex)
        {
            var hexColor = rgbColorModelHex.Val.ToString();
            var hexColorInt = int.Parse(hexColor, NumberStyles.HexNumber, CultureInfo.CurrentCulture);
            Color clr = Color.FromArgb(hexColorInt);

            return new ShapeFill(clr);
        }

        internal static ShapeFill FromASchemeClr(A.SchemeColor schemeColor)
        {
            return new ShapeFill(schemeColor);
        }
        
        internal static ShapeFill FromImage(SCImage image)
        {
            return new ShapeFill(image);
        }
        
        internal static ShapeFill CreateNoFill()
        {
            return new ShapeFill();
        }
    }
}