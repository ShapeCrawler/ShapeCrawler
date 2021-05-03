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
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    internal class LayoutAutoShape : LayoutShape, IAutoShape, IFontDataReader, ITextBoxContainer // TODO: IFontDataReader is needed?
    {
        private readonly SCImageFactory imageFactory = new ();
        private readonly ResettableLazy<Dictionary<int, FontData>> lvlToFontData;
        private readonly Lazy<ShapeFill> shapeFill;
        private readonly Lazy<SCTextBox> textBox;

        internal LayoutAutoShape(SCSlideLayout slideLayout, P.Shape pShape)
            : base(slideLayout, pShape)
        {
            this.textBox = new Lazy<SCTextBox>(this.GetTextBox);
            this.shapeFill = new Lazy<ShapeFill>(this.TryGetFill);
            this.lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(this.GetLvlToFontData);
        }

        #region Public Properties

        public ITextBox TextBox => this.textBox.Value; // TODO: add test

        public ShapeFill Fill => this.shapeFill.Value; // TODO: add test

        public Shape ParentShape { get; }

        #endregion Public Properties

        internal Dictionary<int, FontData> LvlToFontData => this.lvlToFontData.Value;

        internal ShapeContext Context { get; } // TODO: resolve warning

        public void FillFontData(int paragraphLvl, ref FontData fontData)
        {
            // Tries get font from Auto Shape
            if (this.LvlToFontData.TryGetValue(paragraphLvl, out FontData layoutFontData))
            {
                fontData = layoutFontData;
                if (!fontData.IsFilled() && this.Placeholder != null)
                {
                    Placeholder placeholder = (Placeholder)this.Placeholder;
                    IFontDataReader referencedMasterShape = (IFontDataReader) placeholder.ReferencedShape;
                    referencedMasterShape?.FillFontData(paragraphLvl, ref fontData);
                }

                return;
            }

            if (this.Placeholder != null)
            {
                Placeholder placeholder = (Placeholder) this.Placeholder;
                IFontDataReader referencedMasterShape = (IFontDataReader) placeholder.ReferencedShape;
                if (referencedMasterShape != null)
                {
                    referencedMasterShape.FillFontData(paragraphLvl, ref fontData);
                }
            }
        }

        private Dictionary<int, FontData> GetLvlToFontData()
        {
            P.Shape pShape = (P.Shape)this.PShapeTreeChild;
            Dictionary<int, FontData> lvlToFontData = FontDataParser.FromCompositeElement(pShape.TextBody.ListStyle);

            // TODO: move this block to FontDataParser.FromCompositeElement()?
            if (!lvlToFontData.Any())
            {
                Int32Value endParaRunPrFs = pShape.TextBody.GetFirstChild<A.Paragraph>()
                    .GetFirstChild<A.EndParagraphRunProperties>()?.FontSize;
                if (endParaRunPrFs != null)
                {
                    var fontData = new FontData
                    {
                        FontSize = endParaRunPrFs
                    };
                    lvlToFontData.Add(1, fontData);
                }
            }

            return lvlToFontData;
        }

        private SCTextBox GetTextBox()
        {
            P.TextBody pTextBody = this.PShapeTreeChild.GetFirstChild<P.TextBody>();
            if (pTextBody == null)
            {
                return null;
            }

            IEnumerable<A.Text> aTexts = pTextBody.Descendants<A.Text>();
            if (aTexts.Sum(t => t.Text.Length) > 0)
            {
                return new SCTextBox(pTextBody, this);
            }

            return null;
        }

        private ShapeFill TryGetFill() // TODO: duplicate of SlideAutoShape.TryGetFill()
        {
            SCImage image = this.imageFactory.FromSlidePart(this.Context.SlidePart, this.PShapeTreeChild);
            if (image != null)
            {
                return new ShapeFill(image);
            }

            A.SolidFill aSolidFill =
                ((P.Shape) this.PShapeTreeChild).ShapeProperties.GetFirstChild<A.SolidFill>(); // <a:solidFill>
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

        public override SCSlideMaster ParentSlideMaster { get; }
    }
}