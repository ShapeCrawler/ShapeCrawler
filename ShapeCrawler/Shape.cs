using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape.
    /// </summary>
    public abstract class Shape
    {
        
        protected Shape(OpenXmlCompositeElement pShapeTreeChild, SlideSc slide)
        {
            PShapeTreeChild = pShapeTreeChild;
            Slide = slide;
            
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

        protected Shape()
        {
        }

        internal OpenXmlCompositeElement PShapeTreeChild { get; }
        internal SlideSc Slide { get; }

        protected void SetCustomData(string value)
        {
            var customDataElement =
                $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
            PShapeTreeChild.InnerXml += customDataElement;
        }

        private string GetCustomData()
        {
            var pattern = @$"<{ConstantStrings.CustomDataElementName}>(.*)<\/{ConstantStrings.CustomDataElementName}>";
            var regex = new Regex(pattern);
            var elementText = regex.Match(PShapeTreeChild.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }

        #region Public Properties

        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }

        /// <summary>
        ///     Gets placeholder. Returns <c>NULL</c> if the shape is not a placeholder.
        /// </summary>
        public abstract IPlaceholder Placeholder { get; }

        #endregion Public Properties
    }
}