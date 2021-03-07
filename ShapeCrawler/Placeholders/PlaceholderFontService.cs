using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable PossibleNullReferenceException

namespace ShapeCrawler.Placeholders
{
    internal class PlaceholderFontService
    {
        private readonly Lazy<HashSet<PlaceholderFontData>> _layoutPlaceholders;
        private readonly Lazy<Dictionary<int, FontData>> _masterBodyFontHeights;
        private readonly Lazy<HashSet<PlaceholderFontData>> _masterPlaceholders;
        private readonly IPlaceholderService _placeholderService;
        private readonly SlidePart _slidePart;

        #region Public Methods

        public int? GetFontSizeByParagraphLvl(P.Shape pShape, int paragraphLvl)
        {
            PlaceholderData placeholderData = _placeholderService.CreatePlaceholderData(pShape);

            // From slide layout element
            PlaceholderFontData lPlaceholder = _layoutPlaceholders.Value.FirstOrDefault(e => e.Equals(placeholderData));
            if (lPlaceholder != null && lPlaceholder.LvlToFontData.ContainsKey(paragraphLvl))
            {
                if (lPlaceholder.LvlToFontData[paragraphLvl].FontSize != null)
                {
                    return lPlaceholder.LvlToFontData[paragraphLvl].FontSize;
                }
            }

            // From slide master element
            PlaceholderFontData mPlaceholder = _masterPlaceholders.Value.FirstOrDefault(e => e.Equals(placeholderData));
            if (mPlaceholder != null && mPlaceholder.LvlToFontData.ContainsKey(paragraphLvl))
            {
                if (mPlaceholder.LvlToFontData[paragraphLvl].FontSize != null)
                {
                    return mPlaceholder.LvlToFontData[paragraphLvl].FontSize;
                }
            }

            // Title type
            P.TextStyles masterGlobalTextStyle = _slidePart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles;
            if (placeholderData.PlaceholderType == PlaceholderType.Title)
            {
                return masterGlobalTextStyle.TitleStyle.Level1ParagraphProperties
                    .GetFirstChild<A.DefaultRunProperties>().FontSize.Value;
            }

            // Master body type placeholder settings
            if (_masterBodyFontHeights.Value.ContainsKey(paragraphLvl))
            {
                if (_masterBodyFontHeights.Value[paragraphLvl].FontSize != null)
                {
                    return _masterBodyFontHeights.Value[paragraphLvl].FontSize;
                }
            }

            return null;
        }

        #endregion Constructors

        #region Constructors

        public PlaceholderFontService(SlidePart slidePart, IPlaceholderService placeholderService)
        {
            _slidePart = slidePart ?? throw new ArgumentNullException(nameof(slidePart));
            _placeholderService = placeholderService ?? throw new ArgumentNullException(nameof(placeholderService));

            P.CommonSlideData layoutSldData = _slidePart.SlideLayoutPart.SlideLayout.CommonSlideData;
            P.CommonSlideData masterSldData = _slidePart.SlideLayoutPart.SlideMasterPart.SlideMaster.CommonSlideData;
            _layoutPlaceholders = new Lazy<HashSet<PlaceholderFontData>>(() => InitLayoutMaster(layoutSldData));
            _masterPlaceholders = new Lazy<HashSet<PlaceholderFontData>>(() => InitLayoutMaster(masterSldData));
            _masterBodyFontHeights = new Lazy<Dictionary<int, FontData>>(() => InitBodyTypePlaceholder(_slidePart));
        }

        public PlaceholderFontService(SlidePart slidePart)
            : this(slidePart, new PlaceholderService(slidePart.SlideLayoutPart))
        {
        }

        #endregion Constructors

        #region Private Methods

        private HashSet<PlaceholderFontData> InitLayoutMaster(P.CommonSlideData layoutMasterCommonSlideData)
        {
            var fontDataPlaceholders = new HashSet<PlaceholderFontData>();
            foreach (P.Shape pShape in layoutMasterCommonSlideData.ShapeTree.Elements<P.Shape>()
                .Where(e => e.IsPlaceholder()))
            {
                PlaceholderFontData fontDataPlaceholder = FromLayoutMasterElement(pShape);
                fontDataPlaceholders.Add(fontDataPlaceholder);
            }

            return fontDataPlaceholders;
        }

        private static Dictionary<int, FontData> InitBodyTypePlaceholder(SlidePart slidePart)
        {
            P.BodyStyle slideMasterBodyTextStyle =
                slidePart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles.BodyStyle;
            return FontDataParser.FromCompositeElement(slideMasterBodyTextStyle);
        }

        private PlaceholderFontData FromLayoutMasterElement(P.Shape pShape)
        {
            var placeholderFontData = _placeholderService.PlaceholderFontDataFromCompositeElement(pShape);
            placeholderFontData.LvlToFontData = FontDataParser.FromCompositeElement(pShape.TextBody.ListStyle);

            if (!placeholderFontData.LvlToFontData.Any()) // font height is still not known
            {
                var endParaRunPrFs = pShape.TextBody.GetFirstChild<A.Paragraph>()
                    .GetFirstChild<A.EndParagraphRunProperties>()?.FontSize;
                if (endParaRunPrFs != null)
                {
                    placeholderFontData.LvlToFontData.Add(1, new FontData(endParaRunPrFs));
                }
            }

            return placeholderFontData;
        }

        #endregion Private Methods
    }
}