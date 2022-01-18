using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents an AutoShape on a Slide Master.
    /// </summary>
    internal class MasterAutoShape : MasterShape, IAutoShape, ITextBoxContainer, IFontDataReader
    {
        private readonly ResettableLazy<Dictionary<int, FontData>> lvlToFontData;
        private readonly Lazy<ShapeFill> shapeFill;
        private readonly Lazy<SCTextBox?> textBox;
        private readonly P.Shape pShape;

        internal MasterAutoShape(P.Shape pShape, SCSlideMaster parentSlideInternalMaster)
            : base(pShape, parentSlideInternalMaster)
        {
            this.textBox = new Lazy<SCTextBox?>(this.GetTextBox);
            this.shapeFill = new Lazy<ShapeFill>(this.TryGetFill);
            this.lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(this.GetLvlToFontData);
            this.pShape = pShape;
        }

        #region Public Properties

        public ITextBox? TextBox => this.textBox.Value;

        public ShapeFill Fill => this.shapeFill.Value;

        #endregion Public Properties

        internal ShapeContext Context { get; }

        internal Dictionary<int, FontData> LvlToFontData => this.lvlToFontData.Value;

        public void FillFontData(int paragraphLvl, ref FontData fontData)
        {
            if (this.LvlToFontData.TryGetValue(paragraphLvl, out FontData masterFontData) && !fontData.IsFilled())
            {
                masterFontData.Fill(fontData);
                return;
            }

            P.TextStyles pTextStyles = ((SCSlideMaster)this.ParentSlideMaster).PSlideMaster.TextStyles;
            if (this.Placeholder.Type == PlaceholderType.Title)
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

        private SCTextBox GetTextBox() // TODO: duplicate code in LayoutAutoShape
        {
            P.TextBody pTextBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
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

        private ShapeFill TryGetFill() // TODO: duplicate code in LayoutAutoShape
        {
            SCImage image = SCImage.GetFillImageOrDefault(this, this.Context.SlidePart, this.Context.CompositeElement);
            if (image != null)
            {
                return new ShapeFill(image);
            }

            A.SolidFill aSolidFill = this.pShape.ShapeProperties.GetFirstChild<A.SolidFill>(); // <a:solidFill>
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

        public IShape Shape { get; }
        public override SCSlideMaster ParentSlideMaster { get; set; }
    }
}