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
    ///     Represents an Auto Shape on a Slide.
    /// </summary>
    internal class SlideAutoShape : SlideShape, IAutoShape
    {
        private readonly ImageExFactory _imageFactory = new ImageExFactory();
        private readonly ILocation _innerTransform;
        private readonly Lazy<ShapeFill> _shapeFill;
        private readonly Lazy<SCTextBox> _textBox;

        #region Constructors

        internal SlideAutoShape(
            ILocation innerTransform,
            ShapeContext spContext,
            P.Shape pShape,
            SCSlide slide) : base(slide, pShape)
        {
            _innerTransform = innerTransform;
            Context = spContext;
            _textBox = new Lazy<SCTextBox>(GetTextBox);
            _shapeFill = new Lazy<ShapeFill>(TryGetFill);
        }

        #endregion Constructors

        internal ShapeContext Context { get; }

        #region Public Properties

        public long X
        {
            get => _innerTransform.X;
            set => _innerTransform.SetX(value);
        }

        public long Y
        {
            get => _innerTransform.Y;
            set => _innerTransform.SetY(value);
        }

        public long Width
        {
            get => _innerTransform.Width;
            set => _innerTransform.SetWidth(value);
        }

        public long Height
        {
            get => _innerTransform.Height;
            set => _innerTransform.SetHeight(value);
        }

        public ITextBox TextBox => _textBox.Value;

        public ShapeFill Fill => _shapeFill.Value;

        #endregion Properties

        #region Private Methods

        private SCTextBox GetTextBox()
        {
            P.TextBody pTextBody = PShapeTreeChild.GetFirstChild<P.TextBody>();
            if (pTextBody == null)
            {
                return null;
            }

            var aTexts = pTextBody.Descendants<A.Text>();
            if (aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element with text must be exist
            {
                return new SCTextBox(pTextBody, this);
            }

            return null;
        }

        private ShapeFill TryGetFill() //TODO: duplicate of LayoutAutoShape.TryGetFill()
        {
            SCImage image = _imageFactory.TryFromSdkShape(Slide.SlidePart, PShapeTreeChild);
            if (image != null)
            {
                return new ShapeFill(image);
            }

            A.SolidFill aSolidFill =
                ((P.Shape) PShapeTreeChild).ShapeProperties.GetFirstChild<A.SolidFill>(); // <a:solidFill>
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

        #endregion Private Methods
    }
}