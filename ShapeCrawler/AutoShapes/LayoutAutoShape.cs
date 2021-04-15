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
    /// <summary>
    ///     Represents an Auto Shape on a Slide Layout.
    /// </summary>
    internal class LayoutAutoShape : LayoutShape, IAutoShape, IFontDataReader //TODO: IFontDataReader is needed?
    {
        private readonly SCImageFactory _imageFactory = new();
        private readonly ResettableLazy<Dictionary<int, FontData>> _lvlToFontData;
        private readonly Lazy<ShapeFill> _shapeFill;
        private readonly Lazy<SCTextBox> _textBox;

        #region Constructors

        internal LayoutAutoShape(SCSlideLayout slideLayout, P.Shape pShape) : base(slideLayout, pShape)
        {
            _textBox = new Lazy<SCTextBox>(GetTextBox);
            _shapeFill = new Lazy<ShapeFill>(TryGetFill);
            _lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(GetLvlToFontData);
        }

        #endregion Constructors

        internal Dictionary<int, FontData> LvlToFontData => _lvlToFontData.Value;
        internal ShapeContext Context { get; } //TODO: resolve warning

        public void FillFontData(int paragraphLvl, ref FontData fontData)
        {
            // Tries get font from Auto Shape
            if (LvlToFontData.TryGetValue(paragraphLvl, out FontData layoutFontData))
            {
                fontData = layoutFontData;
                if (!fontData.IsFilled() && Placeholder != null)
                {
                    Placeholder placeholder = (Placeholder) Placeholder;
                    IFontDataReader referencedMasterShape = (IFontDataReader) placeholder.ReferencedShape;
                    if (referencedMasterShape != null)
                    {
                        referencedMasterShape.FillFontData(paragraphLvl, ref fontData);
                    }
                }

                return;
            }

            if (Placeholder != null)
            {
                Placeholder placeholder = (Placeholder) Placeholder;
                IFontDataReader referencedMasterShape = (IFontDataReader) placeholder.ReferencedShape;
                if (referencedMasterShape != null)
                {
                    referencedMasterShape.FillFontData(paragraphLvl, ref fontData);
                }
            }
        }

        #region Public Properties

        public ITextBox TextBox => _textBox.Value; // TODO: add test

        public ShapeFill Fill => _shapeFill.Value; // TODO: add test

        #endregion Public Properties

        #region Private Methods

        private Dictionary<int, FontData> GetLvlToFontData()
        {
            P.Shape pShape = (P.Shape) PShapeTreeChild;
            Dictionary<int, FontData> lvlToFontData = FontDataParser.FromCompositeElement(pShape.TextBody.ListStyle);

            if (!lvlToFontData.Any()) // TODO: move this block to FontDataParser.FromCompositeElement()?
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

        private ShapeFill TryGetFill() //TODO: duplicate of SlideAutoShape.TryGetFill()
        {
            SCImage image = _imageFactory.FromSlidePart(Context.SlidePart, PShapeTreeChild);
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