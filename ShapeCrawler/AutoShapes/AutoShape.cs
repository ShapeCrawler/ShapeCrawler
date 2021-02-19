using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories.Drawing;
using ShapeCrawler.Factories.Placeholders;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Models.Styles;
using ShapeCrawler.Settings;
using ShapeCrawler.Statics;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <inheritdoc cref="IAutoShape" />
    public class AutoShape : IAutoShape
    {
        #region Constructors

        internal AutoShape(
            ILocation innerTransform,
            ShapeContext spContext,
            GeometryType geometryType,
            OpenXmlCompositeElement shapeTreeSource,
            SlideSc slide)
        {
            _innerTransform = innerTransform;
            Context = spContext;
            _textBox = new Lazy<TextBoxSc>(GetTextBox);
            _shapeFill = new Lazy<ShapeFill>(TryGetFill);
            ShapeTreeSource = shapeTreeSource;
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
        internal OpenXmlCompositeElement ShapeTreeSource { get; }
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

        public int Id
        {
            get
            {
                InitIdHiddenName();
                return _id;
            }
        }

        public string Name
        {
            get
            {
                InitIdHiddenName();
                return _name;
            }
        }

        public bool Hidden
        {
            get
            {
                InitIdHiddenName();
                return (bool) _hidden;
            }
        }

        public TextBoxSc TextBox => _textBox.Value;

        public Placeholder Placeholder
        {
            get
            {
                if (Context.CompositeElement.IsPlaceholder())
                {
                    return new Placeholder(ShapeTreeSource);
                }

                return null;
            }
        }

        public ShapeFill Fill => _shapeFill.Value;

        public GeometryType GeometryType { get; }

        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }

        #endregion Properties

        #region Private Methods

        private void SetCustomData(string value)
        {
            var customDataElement =
                $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
            Context.CompositeElement.InnerXml += customDataElement;
        }

        private string GetCustomData()
        {
            var pattern = @$"<{ConstantStrings.CustomDataElementName}>(.*)<\/{ConstantStrings.CustomDataElementName}>";
            var regex = new Regex(pattern);
            var elementText = regex.Match(Context.CompositeElement.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }

        private TextBoxSc GetTextBox()
        {
            // TODO: try optimize by avoiding Descendants() using.
            P.TextBody pTextBody = Context.CompositeElement.Descendants<P.TextBody>().SingleOrDefault();
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

            P.Shape pShape = (P.Shape) Context.CompositeElement;
            A.SolidFill aSolidFill = pShape.ShapeProperties.GetFirstChild<A.SolidFill>(); // <a:solidFill>
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