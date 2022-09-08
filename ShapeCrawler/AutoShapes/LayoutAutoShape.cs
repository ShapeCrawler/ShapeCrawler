using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    internal class LayoutAutoShape : LayoutShape, IAutoShape, IFontDataReader, ITextBoxContainer
    {
        private readonly ResettableLazy<Dictionary<int, FontData>> lvlToFontData;
        private readonly Lazy<ShapeFill> shapeFill;
        private readonly Lazy<SCTextBox?> textBox;
        private readonly P.Shape _childOfPShape;

        internal LayoutAutoShape(SCSlideLayout slideLayout, P.Shape childOfPShape)
            : base(slideLayout, childOfPShape)
        {
            this.textBox = new Lazy<SCTextBox?>(this.GetTextBox);
            this.shapeFill = new Lazy<ShapeFill>(TryGetFill);
            this.lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(this.GetLvlToFontData);
            this._childOfPShape = childOfPShape;
        }

        #region Public Properties

        public ITextBox? TextBox => this.textBox.Value;

        public IShapeFill Fill => this.shapeFill.Value;

        public ShapeType ShapeType => ShapeType.AutoShape;
        #endregion Public Properties

        public IShape Shape => this;

        private Dictionary<int, FontData> LvlToFontData => this.lvlToFontData.Value;

        public void FillFontData(int paragraphLvl, ref FontData fontData)
        {
            // Tries get font from Auto Shape
            if (this.LvlToFontData.TryGetValue(paragraphLvl, out FontData layoutFontData))
            {
                fontData = layoutFontData;
                if (!fontData.IsFilled() && this.Placeholder != null)
                {
                    Placeholder placeholder = (Placeholder)this.Placeholder;
                    IFontDataReader referencedMasterShape = (IFontDataReader)placeholder.ReferencedShape;
                    referencedMasterShape?.FillFontData(paragraphLvl, ref fontData);
                }

                return;
            }

            if (this.Placeholder != null)
            {
                Placeholder placeholder = (Placeholder)this.Placeholder;
                IFontDataReader referencedMasterShape = (IFontDataReader)placeholder.ReferencedShape;
                if (referencedMasterShape != null)
                {
                    referencedMasterShape.FillFontData(paragraphLvl, ref fontData);
                }
            }
        }

        private Dictionary<int, FontData> GetLvlToFontData()
        {
            Dictionary<int, FontData> lvlToFontData = FontDataParser.FromCompositeElement(this._childOfPShape.TextBody.ListStyle);

            if (!lvlToFontData.Any())
            {
                Int32Value endParaRunPrFs = this._childOfPShape.TextBody.GetFirstChild<A.Paragraph>()
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
            P.TextBody pTextBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
            if (pTextBody == null)
            {
                return null;
            }

            IEnumerable<A.Text> aTexts = pTextBody.Descendants<A.Text>();
            if (aTexts.Sum(t => t.Text.Length) > 0)
            {
                return new SCTextBox(this, pTextBody);
            }

            return null;
        }

        private static ShapeFill TryGetFill() // TODO: duplicate of SlideAutoShape.TryGetFill()
        {
            throw new NotImplementedException();
        }
    }
}