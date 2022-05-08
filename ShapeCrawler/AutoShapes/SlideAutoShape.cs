using System;
using System.Linq;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    internal class SlideAutoShape : SlideShape, IAutoShape, ITextBoxContainer
    {
        private readonly Lazy<ShapeFill> shapeFill;
        private readonly Lazy<SCTextBox?> textBox;
        private readonly P.Shape pShape;

        public SlideAutoShape(P.Shape pShape, SCSlide parentSlideInternal, SlideGroupShape groupShape)
            : base(pShape, parentSlideInternal, groupShape)
        {
            this.textBox = new Lazy<SCTextBox?>(this.GetTextBox);
            this.shapeFill = new Lazy<ShapeFill>(this.TryGetFill);
            this.pShape = pShape;
        }

        #region Public Properties

        public ITextBox? TextBox => this.textBox.Value;

        public ShapeFill Fill => this.shapeFill.Value;

        public IShape Shape => this; // TODO: should be internal?

        #endregion Public Properties

        private SCTextBox? GetTextBox()
        {
            var pTextBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
            if (pTextBody == null)
            {
                return null;
            }

            var aTexts = pTextBody.Descendants<A.Text>();
            if (aTexts.Sum(t => t.Text.Length) > 0)
            {
                return new SCTextBox(pTextBody, this);
            }

            return null;
        }

        private ShapeFill TryGetFill() // TODO: duplicate of LayoutAutoShape.TryGetFill()
        {
            var slide = (SCSlide)this.Slide;
            SCImage image = SCImage.GetFillImageOrDefault(this, slide.SDKSlidePart, this.PShapeTreesChild);

            if (image != null)
            {
                return new ShapeFill(image);
            }

            A.SolidFill aSolidFill = this.pShape.ShapeProperties.GetFirstChild<A.SolidFill>(); // <a:solidFill>
            if (aSolidFill == null)
            {
                return null;
            }

            A.RgbColorModelHex aRgbColorModelHex = aSolidFill.RgbColorModelHex;
            if (aRgbColorModelHex != null)
            {
                return ShapeFill.FromXmlSolidFill(aRgbColorModelHex);
            }

            return ShapeFill.FromASchemeClr(aSolidFill.SchemeColor);
        }

        public ShapeType ShapeType => ShapeType.AutoShape;
    }
}