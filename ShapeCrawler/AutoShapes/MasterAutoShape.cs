using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
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
    internal class MasterAutoShape : MasterShape, IAutoShape, IAutoShapeInternal
    {
        private readonly ResettableLazy<Dictionary<int, FontData>> _lvlToFontData;

        #region Constructors

        internal MasterAutoShape(SlideMasterSc slideMaster, P.Shape pShape) : base(slideMaster, pShape)
        {
            _textBox = new Lazy<SCTextBox>(GetTextBox);
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

            // Title type
            P.TextStyles pTextStyles = SlideMaster.PSlideMaster.TextStyles;
            if (Placeholder.Type == PlaceholderType.Title)
            {
                var fontSize = pTextStyles.TitleStyle.Level1ParagraphProperties
                    .GetFirstChild<A.DefaultRunProperties>().FontSize.Value;
                fontData = new FontData(fontSize);
                return true;
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

        private readonly Lazy<SCTextBox> _textBox;
        private readonly Lazy<ShapeFill> _shapeFill;
        private readonly ImageExFactory _imageFactory = new ImageExFactory();

        internal ShapeContext Context { get; }

        #endregion Fields

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
                return new SCTextBox(pTextBody, this);
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

        #endregion
    }
}