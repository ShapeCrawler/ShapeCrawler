using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    internal class LayoutAutoShape : LayoutShape, IAutoShape, IFontDataReader, ITextFrameContainer
    {
        private readonly ResettableLazy<Dictionary<int, FontData>> lvlToFontData;
        private readonly Lazy<ShapeFill> shapeFill;
        private readonly Lazy<TextFrame?> textBox;
        private readonly P.Shape pShape;

        internal LayoutAutoShape(SCSlideLayout slideLayout, P.Shape pShape)
            : base(slideLayout, pShape)
        {
            this.textBox = new Lazy<TextFrame?>(this.GetTextFrame);
            this.shapeFill = new Lazy<ShapeFill>(TryGetFill);
            this.lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(this.GetLvlToFontData);
            this.pShape = pShape;
        }

        #region Public Properties

        public ITextFrame? TextFrame => this.textBox.Value;

        public IShapeFill Fill => this.shapeFill.Value;

        public SCShapeType ShapeType => SCShapeType.AutoShape;
        
        #endregion Public Properties

        public Shape Shape => this;

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
            Dictionary<int, FontData> lvlToFontData = FontDataParser.FromCompositeElement(this.pShape.TextBody.ListStyle);

            if (!lvlToFontData.Any())
            {
                Int32Value endParaRunPrFs = this.pShape.TextBody.GetFirstChild<A.Paragraph>()
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

        private TextFrame GetTextFrame()
        {
            P.TextBody pTextBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
            if (pTextBody == null)
            {
                return null;
            }

            IEnumerable<A.Text> aTexts = pTextBody.Descendants<A.Text>();
            if (aTexts.Sum(t => t.Text.Length) > 0)
            {
                return new TextFrame(this, pTextBody, false);
            }

            return null;
        }

        private static ShapeFill TryGetFill() // TODO: duplicate of SlideAutoShape.TryGetFill()
        {
            throw new NotImplementedException();
        }
    }
}