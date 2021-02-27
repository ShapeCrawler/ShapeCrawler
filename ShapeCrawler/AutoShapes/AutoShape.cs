using System;
using System.Linq;
using System.Text.RegularExpressions;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Settings;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <inheritdoc cref="IAutoShape" />
    public class AutoShape : Shape, IAutoShape
    {
        #region Constructors

        internal AutoShape(
            ILocation innerTransform,
            ShapeContext spContext,
            GeometryType geometryType,
            P.Shape pShape,
            SlideSc slide) : base(pShape)
        {
            _innerTransform = innerTransform;
            Context = spContext;
            _textBox = new Lazy<TextBoxSc>(GetTextBox);
            _shapeFill = new Lazy<ShapeFill>(TryGetFill);
            GeometryType = geometryType;
            Slide = slide;
        }

        #endregion Constructors

        #region Fields

        private readonly Lazy<TextBoxSc> _textBox;
        private readonly Lazy<ShapeFill> _shapeFill;
        private readonly ImageExFactory _imageFactory = new ImageExFactory();
        private bool? _hidden;
        private int _id;
        private string _name;
        private readonly ILocation _innerTransform;

        internal ShapeContext Context { get; }
        internal SlideSc Slide { get; }

        #endregion Fields

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

        public int Id //TODO: move to Shape
        {
            get
            {
                InitIdHiddenName();
                return _id;
            }
        }

        public string Name //TODO: move to Shape
        {
            get
            {
                InitIdHiddenName();
                return _name;
            }
        }

        public bool Hidden //TODO: move to Shape
        {
            get
            {
                InitIdHiddenName();
                return (bool) _hidden;
            }
        }

        public ITextBox TextBox => _textBox.Value;

        public ShapeFill Fill => _shapeFill.Value;

        public GeometryType GeometryType { get; }

        #endregion Properties

        #region Private Methods

        private TextBoxSc GetTextBox()
        {
            P.TextBody pTextBody = PShapeTreeChild.GetFirstChild<P.TextBody>();
            if (pTextBody == null)
            {
                return null;
            }

            var aTexts = pTextBody.Descendants<A.Text>();
            if (aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element with text must be exist
            {
                return new TextBoxSc(this, pTextBody);
            }

            return null;
        }

        private ShapeFill TryGetFill()
        {
            ImageSc image = _imageFactory.TryFromSdkShape(Context.SlidePart, Context.CompositeElement);
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

        private void InitIdHiddenName()
        {
            if (_id != 0)
            {
                return;
            }

            var (id, hidden, name) = Context.CompositeElement.GetNvPrValues();
            _id = id;
            _hidden = hidden;
            _name = name;
        }

        #endregion
    }
}