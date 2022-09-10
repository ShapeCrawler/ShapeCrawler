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
        private bool isInitialized = false;
        private FillType fillType;

        internal ShapeFill(Shape shape)
        {
            this.shape = shape;
        }

        public FillType Type => this.GetFillType();

        private FillType GetFillType()
        {
            if (!this.isInitialized)
            {
                this.Initialize();
            }

            return this.fillType;
        }

        private void Initialize()
        {
            var xmlPart = this.shape.SlideBase.TypedOpenXmlPart;
            var image = SCImage.ForAutoShapeFill(shape, xmlPart);

            if (image != null)
            {
                this.Picture = image;
                this.fillType = FillType.Picture;
                return;
            }

            var pShape = (P.Shape)this.shape.PShapeTreesChild;
            var aSolidFill = pShape.ShapeProperties!.GetFirstChild<A.SolidFill>(); 
            if (aSolidFill == null)
            {
                this.fillType = FillType.NoFill;
                return;
            }

            var aRgbColorModelHex = aSolidFill.RgbColorModelHex;
            if (aRgbColorModelHex != null)
            {
                var hexColor = aRgbColorModelHex.Val!.ToString();
                this.HexSolidColor = hexColor;
                
                this.fillType = FillType.Solid;
                return;
            }

            var schemeColor = aSolidFill.SchemeColor;
            this.fillType = FillType.Picture;
        }

        public SCImage? Picture { get; private set; }
        public string HexSolidColor { get; private set; }

        public Color SolidColor { get; }

        public void SetPicture(Stream imageStream)
        {
            if (this.Type == FillType.NoFill)
            {
                var rId = this.shape.SlideBase.TypedOpenXmlPart.AddImagePart(imageStream);

                var aBlipFill = new A.BlipFill();
                var aBlip = new A.Blip { Embed = rId };
                var stretch = new A.Stretch();
                var fillRectangle = new A.FillRectangle();
                stretch.Append(fillRectangle);
                aBlipFill.Append(aBlip);
                aBlipFill.Append(stretch);

                this.shape.PShapeTreesChild.GetFirstChild<P.ShapeProperties>() !.Append(aBlipFill);
            }

            this.isInitialized = false;
        }
    }
}