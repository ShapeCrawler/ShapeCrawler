using System;
using System.Linq;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Factories;
using ShapeCrawler.Settings;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents an AutoShape on a slide.
    /// </summary>
    internal class SlideAutoShape : SlideShape, IAutoShape, ITextBoxContainer
    {
        private readonly SCImageFactory imageFactory = new (); // TODO: make it static call
        private readonly ILocation innerTransform;
        private readonly Lazy<ShapeFill> shapeFill;
        private readonly Lazy<SCTextBox> textBox;
        private readonly P.Shape pShape;

        internal SlideAutoShape(
            ILocation innerTransform,
            ShapeContext spContext,
            P.Shape pShape,
            SCSlide slide)
            : base(slide, pShape)
        {
            this.innerTransform = innerTransform;
            this.Context = spContext;
            this.textBox = new Lazy<SCTextBox>(this.GetTextBox);
            this.shapeFill = new Lazy<ShapeFill>(this.TryGetFill);
            this.pShape = pShape;
        }

        internal ShapeContext Context { get; }

        #region Public Properties

        public long X // TODO: remove this hides
        {
            get => this.innerTransform.X;
            set => this.innerTransform.SetX(value);
        }

        public long Y
        {
            get => this.innerTransform.Y;
            set => this.innerTransform.SetY(value);
        }

        public long Width
        {
            get => this.innerTransform.Width;
            set => this.innerTransform.SetWidth(value);
        }

        public long Height
        {
            get => this.innerTransform.Height;
            set => this.innerTransform.SetHeight(value);
        }

        public ITextBox TextBox => this.textBox.Value;

        public ShapeFill Fill => this.shapeFill.Value;

        #endregion Properties

        private SCTextBox GetTextBox()
        {
            P.TextBody pTextBody = this.PShapeTreeChild.GetFirstChild<P.TextBody>();
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
            SCImage image = this.imageFactory.FromSlidePart(this.Slide.SlidePart, this.PShapeTreeChild);
            if (image != null)
            {
                return new ShapeFill(image);
            }

            A.SolidFill aSolidFill = this.pShape.ShapeProperties.GetFirstChild<A.SolidFill>(); // <a:solidFill>
            if (aSolidFill != null)
            {
                A.RgbColorModelHex aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    return ShapeFill.FromXmlSolidFill(aRgbColorModelHex);
                }

                return ShapeFill.FromASchemeClr(aSolidFill.SchemeColor);
            }

            return null;
        }
    }
}