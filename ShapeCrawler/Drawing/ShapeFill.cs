using System.Drawing;
using System.Globalization;
using System.IO;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing
{
    internal class ShapeFill : IShapeFill
    {
        private readonly Shape shape;

        #region Contructors

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

        public FillType Type { get; private set; }

        public SCImage? Picture { get; }

        public Color SolidColor { get; }

        public void SetPicture(Stream imageStream)
        {
            if (this.Type == FillType.NoFill)
            {
                var rId = this.shape.SlideBase.OpenXmlPart.AddImagePart(imageStream);

                var aBlipFill = new A.BlipFill();
                var aBlip = new A.Blip { Embed = rId };
                var stretch = new A.Stretch();
                var fillRectangle = new A.FillRectangle();
                stretch.Append(fillRectangle);
                aBlipFill.Append(aBlip);
                aBlipFill.Append(stretch);

                this.shape.PShapeTreesChild.GetFirstChild<P.ShapeProperties>() !.Append(aBlipFill);
            }

            this.Type = FillType.Picture;
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