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
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents an Auto Shape on a Slide.
    /// </summary>
    internal class SlideAutoShape : SlideShape, IAutoShape, IAutoShapeInternal
    {
        private readonly ImageExFactory _imageFactory = new ImageExFactory();
        private readonly ILocation _innerTransform;
        private readonly ResettableLazy<Dictionary<int, FontData>> _lvlToFontData;
        private readonly Lazy<ShapeFill> _shapeFill;
        private readonly Lazy<SCTextBox> _textBox;
        private bool? _hidden;
        private int _id;
        private string _name;

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
            _lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(() => GetLvlToFontData());
        }

        #endregion Constructors

        internal ShapeContext Context { get; }

        internal Dictionary<int, FontData> LvlToFontData => _lvlToFontData.Value;

        public bool TryGetFontData(int paragraphLvl, out FontData fontData)
        {
            // Tries get font from Auto Shape
            if (LvlToFontData.TryGetValue(paragraphLvl, out fontData))
            {
                return true;
            }

            // Tries get font from Auto Shape of Placeholder
            if (Placeholder == null)
            {
                return false;
            }

            Placeholder placeholder = (Placeholder) Placeholder;
            IAutoShapeInternal placeholderAutoShape = (IAutoShapeInternal) placeholder.Shape;
            return placeholderAutoShape.TryGetFontData(paragraphLvl, out fontData);
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

        #region Public Properties

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
                return new SCTextBox(this, pTextBody);
            }

            return null;
        }

        private ShapeFill TryGetFill()
        {
            SCImage image = _imageFactory.TryFromSdkShape(Context.SlidePart, Context.CompositeElement);
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