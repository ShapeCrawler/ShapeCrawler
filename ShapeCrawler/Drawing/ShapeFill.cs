using System.Drawing;
using System.Globalization;
using System.IO;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing
{
    internal class ShapeFill : IShapeFill
    {
        public ShapeFill()
        {
            this.Type = FillType.NoFill;
        }
        
        public ShapeFill(SCImage image)
        {
            this.Picture = image;
            this.Type = FillType.Picture;
        }
        
        private ShapeFill(Color clr)
        {
            this.SolidColor = clr;
            this.Type = FillType.Solid;
        }

        private ShapeFill(A.SchemeColor schemeColor)
        {
        }
        
        public FillType Type { get; }

        public SCImage? Picture { get; }

        public Color SolidColor { get; }

        public static ShapeFill FromXmlSolidFill(A.RgbColorModelHex rgbColorModelHex)
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

        public void SetPicture(MemoryStream inputPngStream)
        {
            if (this.Type == FillType.NoFill)
            {
                
            }
        }
    }
}