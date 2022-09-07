using System.Drawing;
using System.Globalization;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing
{
    internal class ShapeFill : IShapeFill
    {
        private Shape shape;
        
        #region Contrusctors

        private ShapeFill(Shape shape)
        {
            this.shape = shape;
            this.Type = FillType.NoFill;
        }
        
        private ShapeFill(Shape shape, SCImage image)
        {
            this.shape = shape;
            this.Picture = image;
            this.Type = FillType.Picture;
        }
        
        private ShapeFill(Shape shape, Color color)
        {
            this.shape = shape;
            this.SolidColor = color;
            this.Type = FillType.Solid;
        }

        private ShapeFill(Shape shape, A.SchemeColor schemeColor)
        {
            this.shape = shape;
        }
        
        #endregion
        
        public FillType Type { get; }

        public SCImage? Picture { get; }

        public Color SolidColor { get; }
        
        public void SetPicture(Stream image)
        {
            if (this.Type == FillType.NoFill)
            {
                var aBlipFill = new A.BlipFill();
                var aBlip = new A.Blip();
            }
        }

        internal static ShapeFill WithHexColor(Shape shape, A.RgbColorModelHex rgbColorModelHex)
        {
            var hexColor = rgbColorModelHex.Val!.ToString();
            var hexColorInt = int.Parse(hexColor, NumberStyles.HexNumber, CultureInfo.CurrentCulture);
            Color clr = Color.FromArgb(hexColorInt);

            return new ShapeFill(shape, clr);
        }

        internal static ShapeFill WithSchemeColor(Shape shape, A.SchemeColor schemeColor)
        {
            return new ShapeFill(shape, schemeColor);
        }
        
        internal static ShapeFill WithPicture(Shape shape, SCImage image)
        {
            return new ShapeFill(shape, image);
        }
        
        internal static ShapeFill WithNoFill(Shape shape)
        {
            return new ShapeFill(shape);
        }
    }
}