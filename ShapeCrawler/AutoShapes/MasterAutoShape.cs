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
    /// <summary>
    ///     Represents an AutoShape on a Slide Master.
    /// </summary>
    internal class MasterAutoShape : MasterShape, IAutoShape, ITextBoxContainer, IFontDataReader
    {
        private readonly SCImageFactory imageFactory = new ();
        private readonly ResettableLazy<Dictionary<int, FontData>> lvlToFontData;
        private readonly Lazy<ShapeFill> shapeFill;
        private readonly Lazy<SCTextBox> textBox;

        internal MasterAutoShape(SCSlideMaster slideMaster, P.Shape pShape)
            : base(slideMaster, pShape)
        {
            this.textBox = new Lazy<SCTextBox>(this.GetTextBox);
            this.shapeFill = new Lazy<ShapeFill>(this.TryGetFill);
            this.lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(this.GetLvlToFontData);
        }

        internal ShapeContext Context { get; }

        internal Dictionary<int, FontData> LvlToFontData => this.lvlToFontData.Value;

        public void FillFontData(int paragraphLvl, ref FontData fontData)
        {
            if (this.LvlToFontData.TryGetValue(paragraphLvl, out FontData masterFontData) && !fontData.IsFilled())
            {
                masterFontData.Fill(fontData);
                return;
            }

            P.TextStyles pTextStyles = this.SlideMaster.PSlideMaster.TextStyles;
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
            P.Shape pShape = (P.Shape) this.PShapeTreeChild;
            Dictionary<int, FontData> lvlToFontData = FontDataParser.FromCompositeElement(pShape.TextBody.ListStyle);

            if (!lvlToFontData.Any())
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

        private SCTextBox GetTextBox() // TODO: duplicate code in LayoutAutoShape
        {
            P.TextBody pTextBody = PShapeTreeChild.GetFirstChild<P.TextBody>();
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

        private ShapeFill TryGetFill() //TODO: duplicate code in LayoutAutoShape
        {
            SCImage image = imageFactory.FromSlidePart(Context.SlidePart, Context.CompositeElement);
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

        public Shape ParentShape { get; }

        #region Public Properties

        public ITextBox TextBox => textBox.Value;

        public ShapeFill Fill => shapeFill.Value;

        #endregion Public Properties
    }
}