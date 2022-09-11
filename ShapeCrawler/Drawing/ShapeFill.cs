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
        private string? hexSolidColor;
        private SCImage? pictureImage;
        private A.SolidFill? aSolidFill;
        private A.GradientFill? aGradFill;
        private A.PatternFill? aPattFill;

        internal ShapeFill(Shape shape)
        {
            this.shape = shape;
        }

        public string HexSolidColor => this.GetHexSolidColor();

        private string GetHexSolidColor()
        {
            if (!this.isInitialized)
            {
                this.Initialize();
            }

            return this.hexSolidColor;
        }

        public SCImage? Picture => GetPicture();

        private SCImage GetPicture()
        {
            if (!this.isInitialized)
            {
                this.Initialize();
            }

            return this.pictureImage;
        }

        public Color SolidColor { get; }

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
            var pShape = (P.Shape)this.shape.PShapeTreesChild;
            this.aSolidFill = pShape.ShapeProperties!.GetFirstChild<A.SolidFill>();
            if (this.aSolidFill != null)
            {
                var aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    var hexColor = aRgbColorModelHex.Val!.ToString();
                    this.hexSolidColor = hexColor;
                }
                else
                {
                    // TODO: get hex color from scheme
                    var schemeColor = aSolidFill.SchemeColor;    
                }

                this.fillType = FillType.Solid;
                return;
            }

            this.aGradFill = pShape.ShapeProperties.GetFirstChild<A.GradientFill>();
            if (this.aGradFill != null)
            {
                this.fillType = FillType.Gradient;
                return;
            }
            
            this.aPattFill = pShape.ShapeProperties.GetFirstChild<A.PatternFill>();
            if (this.aPattFill != null)
            {
                this.fillType = FillType.Pattern;
                return;
            }
            
            var xmlPart = this.shape.SlideBase.TypedOpenXmlPart;
            var image = SCImage.ForAutoShapeFill(shape, xmlPart);
            if (image != null)
            {
                this.pictureImage = image;
                this.fillType = FillType.Picture;
                return;
            }

            if (pShape.UseBackgroundFill != null)
            {
                this.fillType = FillType.SlideBackground;
                return;
            }
            
            this.fillType = FillType.NoFill;
        }

        public void SetPicture(Stream imageStream)
        {
            if (this.Type == FillType.Picture)
            {
                this.pictureImage.SetImage(imageStream);
            }
            else
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
        }
    }
}