using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMaster;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <inheritdoc cref="IAutoShape" />
    internal abstract class AutoShape : Shape, IAutoShape
    {
        internal Dictionary<int, FontData> LvlToFontData => _lvlToFontData.Value;
        private readonly ResettableLazy<Dictionary<int, FontData>> _lvlToFontData;

        #region Constructors

        internal AutoShape(
            ILocation innerTransform,
            ShapeContext spContext,
            GeometryType geometryType,
            P.Shape pShape,
            SlideSc slide) : base(pShape, slide)
        {
            _innerTransform = innerTransform;
            Context = spContext;
            _textBox = new Lazy<TextBoxSc>(GetTextBox);
            _shapeFill = new Lazy<ShapeFill>(TryGetFill);
            GeometryType = geometryType;
            _lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(() => GetLvlToFontData());
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

        internal AutoShape(SlideLayoutSc slideLayout, P.Shape pShape) : base(pShape, slideLayout)
        {
        }

        internal ShapeContext Context { get; }

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

        internal bool TryGetFontSizeFromShapeOrPlaceholderShape(int paragraphLvl, out int fontSize)
        {
            // Tries get font from Auto Shape
            if (LvlToFontData.TryGetValue(paragraphLvl, out FontData fontData) && fontData.FontSize != null)
            {
                fontSize = fontData.FontSize;
                return true;
            }

            // Tries get font from Auto Shape of Placeholder
            if (Placeholder == null)
            {
                fontSize = -1;
                return false;
            }
            Placeholder placeholder = (Placeholder)Placeholder;
            AutoShape placeholderAutoShape = (AutoShape)placeholder.Shape;
            return placeholderAutoShape.TryGetFontSizeFromShapeOrPlaceholderShape(paragraphLvl, out fontSize);
        }

        internal Dictionary<int, FontData> GetLvlToFontData()
        {
            P.Shape pShape = (P.Shape)PShapeTreeChild;
            Dictionary<int, FontData> lvlToFontData = FontDataParser.FromCompositeElement(pShape.TextBody.ListStyle);

            if (!lvlToFontData.Any()) // font height is still not known
            {
                Int32Value endParaRunPrFs = pShape.TextBody.GetFirstChild<A.Paragraph>()
                    .GetFirstChild<A.EndParagraphRunProperties>()?.FontSize;
                if (endParaRunPrFs != null)
                {
                    lvlToFontData.Add(1, new FontData(endParaRunPrFs));
                }
            }

            return lvlToFontData;
        }
    }
}