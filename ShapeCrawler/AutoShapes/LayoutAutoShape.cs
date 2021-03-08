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
    /// <summary>
    ///     Represents an Auto Shape on a Slide Layout.
    /// </summary>
    internal class LayoutAutoShape : LayoutShape, IAutoShape, IAutoShapeInternal
    {
        private readonly ResettableLazy<Dictionary<int, FontData>> _lvlToFontData;

        #region Constructors

        internal LayoutAutoShape(SlideLayoutSc slideLayout, P.Shape pShape) : base(slideLayout, pShape)
        {
            _textBox = new Lazy<TextBoxSc>(GetTextBox);
            _shapeFill = new Lazy<ShapeFill>(TryGetFill);
            _lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(() => GetLvlToFontData());
        }

        #endregion Constructors

        internal Dictionary<int, FontData> LvlToFontData => _lvlToFontData.Value;

        public bool TryGetFontData(int paragraphLvl, out FontData fontData)
        {
            // Tries get font from Auto Shape
            if (LvlToFontData.TryGetValue(paragraphLvl, out fontData))
            {
                return true;
            }

            if (Placeholder != null)
            {
                Placeholder placeholder = (Placeholder) Placeholder;
                IAutoShapeInternal placeholderAutoShape = (IAutoShapeInternal) placeholder.Shape;
                if (placeholderAutoShape != null)
                {
                    if (placeholderAutoShape.TryGetFontData(paragraphLvl, out fontData))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        internal Dictionary<int, FontData> GetLvlToFontData()
        {
            P.Shape pShape = (P.Shape) PShapeTreeChild;
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

        #region Fields

        private readonly Lazy<TextBoxSc> _textBox;
        private readonly Lazy<ShapeFill> _shapeFill;
        private readonly ImageExFactory _imageFactory = new ImageExFactory();
        private bool? _hidden;
        private int _id;
        private string _name;
        private P.Shape pShape;
        private readonly ILocation _innerTransform;

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