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
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <inheritdoc cref="IAutoShape" />
    internal class MasterAutoShape : MasterShape, IAutoShape, IFontDataReader
    {
        private readonly ImageExFactory _imageFactory = new ImageExFactory();
        private readonly ResettableLazy<Dictionary<int, FontData>> _lvlToFontData;
        private readonly Lazy<ShapeFill> _shapeFill;
        private readonly Lazy<SCTextBox> _textBox;

        #region Constructors

        internal MasterAutoShape(SCSlideMaster slideMaster, P.Shape pShape) : base(slideMaster, pShape)
        {
            _textBox = new Lazy<SCTextBox>(GetTextBox);
            _shapeFill = new Lazy<ShapeFill>(TryGetFill);
            _lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(GetLvlToFontData);
        }

        #endregion Constructors

        internal ShapeContext Context { get; }
        internal Dictionary<int, FontData> LvlToFontData => _lvlToFontData.Value;

        public void FillFontData(int paragraphLvl, ref FontData fontData)
        {
            if (LvlToFontData.TryGetValue(paragraphLvl, out FontData masterFontData) && !fontData.IsFilled())
            {
                masterFontData.Fill(fontData);
                return;
            }

            P.TextStyles pTextStyles = SlideMaster.PSlideMaster.TextStyles;
            if (Placeholder.Type == PlaceholderType.Title)
            {
                int titleFontSize = pTextStyles.TitleStyle.Level1ParagraphProperties
                    .GetFirstChild<A.DefaultRunProperties>().FontSize.Value;
                if (fontData.FontSize == null)
                {
                    fontData.FontSize = new Int32Value(titleFontSize);
                }
            }
        }

        internal Dictionary<int, FontData> GetLvlToFontData() // TODO: duplicate code in LayoutAutoShape
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

        private SCTextBox GetTextBox() //TODO: duplicate code in LayoutAutoShape
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

        private ShapeFill TryGetFill() //TODO: duplicate code in LayoutAutoShape
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

        #region Public Properties

        public ITextBox TextBox => _textBox.Value;

        public ShapeFill Fill => _shapeFill.Value;

        #endregion Public Properties
    }
}