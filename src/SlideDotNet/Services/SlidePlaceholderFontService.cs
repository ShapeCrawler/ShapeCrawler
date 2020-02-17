using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Enums;
using SlideDotNet.Extensions;
using SlideDotNet.Services.Placeholders;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleNullReferenceException

namespace SlideDotNet.Services
{
    /// <summary>
    /// Represents a font size manager for placeholder elements
    /// </summary>
    public class SlidePlaceholderFontService
    {
        private readonly SlidePart _xmlSldPart;
        private readonly Lazy<HashSet<PlaceholderFontData>> _layoutPlaceholders;
        private readonly Lazy<HashSet<PlaceholderFontData>> _masterPlaceholders;
        private readonly Lazy<Dictionary<int, int>> _masterBodyFontHeights;

        public SlidePlaceholderFontService(SlidePart xmlSldPart)
        {
            _xmlSldPart = xmlSldPart ?? throw new ArgumentNullException(nameof(xmlSldPart));
            var layoutPart = _xmlSldPart.SlideLayoutPart;
            _layoutPlaceholders = new Lazy<HashSet<PlaceholderFontData>>(InitLayoutMaster(layoutPart.SlideLayout.CommonSlideData));
            _masterPlaceholders = new Lazy<HashSet<PlaceholderFontData>>(InitLayoutMaster(layoutPart.SlideMasterPart.SlideMaster.CommonSlideData));
            _masterBodyFontHeights = new Lazy<Dictionary<int, int>>(InitBodyTypePlaceholder(_xmlSldPart));
        }

        /// <summary>
        /// Tries gets font size. Return null if font height is not defined.
        /// </summary>
        /// <param name="xmlSlideShape">Placeholder element.</param>
        /// <param name="pLvl">Paragraph level.</param>
        /// <returns></returns>
        public int? TryGetFontHeight(P.Shape xmlSlideShape, int pLvl)
        {
            var sPlaceholder = FromSlideXmlElement(xmlSlideShape);
            
            // From slide layout element
            var lPlaceholder = _layoutPlaceholders.Value.SingleOrDefault(e => e.Equals(sPlaceholder));
            if (lPlaceholder != null && lPlaceholder.LvlFontHeights.ContainsKey(pLvl))
            {
                return lPlaceholder.LvlFontHeights[pLvl];
            }

            // From slide master element
            var mPlaceholder = _masterPlaceholders.Value.SingleOrDefault(e => e.Equals(sPlaceholder));
            if (mPlaceholder != null && mPlaceholder.LvlFontHeights.ContainsKey(pLvl))
            {
                return mPlaceholder.LvlFontHeights[pLvl];
            }

            // Title type
            var masterGlobalTextStyle = _xmlSldPart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles;
            if (sPlaceholder.PlaceholderType == PlaceholderType.Title)
            {
                return masterGlobalTextStyle.TitleStyle.Level1ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value;
            }

            // Master body type placeholder settings
            if (_masterBodyFontHeights.Value.ContainsKey(pLvl))
            {
                return _masterBodyFontHeights.Value[pLvl];
            }

            return null;
        }

        #region Private Methods

        private HashSet<PlaceholderFontData> InitLayoutMaster(P.CommonSlideData layoutMasterCommonSlideData)
        {
            var placeholders = new HashSet<PlaceholderFontData>();
            foreach (var xmlPh in layoutMasterCommonSlideData.ShapeTree.Elements<P.Shape>().Where(e => e.IsPlaceholder()))
            {
                var ph = FromLayoutMasterElement(xmlPh);
                placeholders.Add(ph);
            }

            return placeholders;
        }

        private static Dictionary<int, int> InitBodyTypePlaceholder(SlidePart xmlSldPart)
        {
            return FontHeightParser.FromCompositeElement(xmlSldPart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles.BodyStyle);
        }

        private static PlaceholderFontData FromSlideXmlElement(P.Shape xmlShape)
        {
            var result = new PlaceholderFontData();
            var ph = xmlShape.Descendants<P.PlaceholderShape>().First();
            var phTypeXml = ph.Type;

            // TYPE
            if (phTypeXml == null)
            {
                result.PlaceholderType = PlaceholderType.Custom;
            }
            else
            {
                // Simple title and centered title placeholders were united
                if (phTypeXml == P.PlaceholderValues.Title || phTypeXml == P.PlaceholderValues.CenteredTitle)
                {
                    result.PlaceholderType = PlaceholderType.Title;
                }
                else
                {
                    result.PlaceholderType = Enum.Parse<PlaceholderType>(phTypeXml.Value.ToString());
                }
            }

            // INDEX
            if (ph.Index != null)
            {
                result.Index = (int)ph.Index.Value;
            }

            return result;
        }

        private static PlaceholderFontData FromLayoutMasterElement(P.Shape xmlShape)
        {
            var ph = FromSlideXmlElement(xmlShape);
            ph.LvlFontHeights = FontHeightParser.FromCompositeElement(xmlShape.TextBody.ListStyle);

            if (!ph.LvlFontHeights.Any()) // font height is still not known
            {
                var endParaRunPrFs = xmlShape.TextBody.GetFirstChild<A.Paragraph>().GetFirstChild<A.EndParagraphRunProperties>()?.FontSize;
                if (endParaRunPrFs != null)
                {
                    ph.LvlFontHeights.Add(1, endParaRunPrFs.Value);
                }
            }

            return ph;
        }

        #endregion Private Methods
    }
}
