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
    /// <summary>
    ///     Represents a font height manager for placeholder elements.
    /// </summary>
    public class PlaceholderFontService
    {
        private readonly Lazy<HashSet<PlaceholderFontData>> _layoutPlaceholders;
        private readonly Lazy<Dictionary<int, int>> _masterBodyFontHeights;
        private readonly Lazy<HashSet<PlaceholderFontData>> _masterPlaceholders;
        private readonly IPlaceholderService _placeholderService;
        private readonly SlidePart _slidePart;

        #region Public Methods

        /// <summary>
        ///     Gets font size. Return null if font size is not defined.
        /// </summary>
        /// <param name="pShape">Placeholder element.</param>
        /// <param name="paragraphLvl">Paragraph level.</param>
        /// <returns></returns>
        public int? GetFontSizeByParagraphLvl(P.Shape pShape, int paragraphLvl)
        {
            PlaceholderData placeholderData = _placeholderService.CreatePlaceholderData(pShape);

            // From slide layout element
            PlaceholderFontData lPlaceholder = _layoutPlaceholders.Value.FirstOrDefault(e => e.Equals(placeholderData));
            if (lPlaceholder != null && lPlaceholder.LvlFontHeights.ContainsKey(paragraphLvl))
            {
                return lPlaceholder.LvlFontHeights[paragraphLvl];
            }

            // From slide master element
            PlaceholderFontData mPlaceholder = _masterPlaceholders.Value.FirstOrDefault(e => e.Equals(placeholderData));
            if (mPlaceholder != null && mPlaceholder.LvlFontHeights.ContainsKey(paragraphLvl))
            {
                return mPlaceholder.LvlFontHeights[paragraphLvl];
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
                return _masterBodyFontHeights.Value[paragraphLvl];
            }

            return null;
        }

        #endregion Public Methods

        #region Constructors

        public PlaceholderFontService(SlidePart sdkSldPart, IPlaceholderService placeholderService)
        {
            _slidePart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _placeholderService = placeholderService ?? throw new ArgumentNullException(nameof(placeholderService));

            P.CommonSlideData layoutSldData = _slidePart.SlideLayoutPart.SlideLayout.CommonSlideData;
            P.CommonSlideData masterSldData = _slidePart.SlideLayoutPart.SlideMasterPart.SlideMaster.CommonSlideData;
            _layoutPlaceholders = new Lazy<HashSet<PlaceholderFontData>>(() => InitLayoutMaster(layoutSldData));
            _masterPlaceholders = new Lazy<HashSet<PlaceholderFontData>>(() => InitLayoutMaster(masterSldData));
            _masterBodyFontHeights = new Lazy<Dictionary<int, int>>(() => InitBodyTypePlaceholder(_slidePart));
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
            foreach (var sdkShape in layoutMasterCommonSlideData.ShapeTree.Elements<P.Shape>()
                .Where(e => e.IsPlaceholder()))
            {
                var fontDataPlaceholder = FromLayoutMasterElement(sdkShape);
                fontDataPlaceholders.Add(fontDataPlaceholder);
            }

            return fontDataPlaceholders;
        }

        private static Dictionary<int, int> InitBodyTypePlaceholder(SlidePart slidePart)
        {
            return FontHeightParser.FromCompositeElement(slidePart.SlideLayoutPart.SlideMasterPart.SlideMaster
                .TextStyles.BodyStyle);
        }

        private PlaceholderFontData FromLayoutMasterElement(P.Shape pShape)
        {
            var placeholderFontData = _placeholderService.PlaceholderFontDataFromCompositeElement(pShape);
            placeholderFontData.LvlFontHeights = FontHeightParser.FromCompositeElement(pShape.TextBody.ListStyle);

            if (!placeholderFontData.LvlFontHeights.Any()) // font height is still not known
            {
                var endParaRunPrFs = pShape.TextBody.GetFirstChild<A.Paragraph>()
                    .GetFirstChild<A.EndParagraphRunProperties>()?.FontSize;
                if (endParaRunPrFs != null)
                {
                    placeholderFontData.LvlFontHeights.Add(1, endParaRunPrFs.Value);
                }
            }

            return placeholderFontData;
        }

        #endregion Private Methods
    }
}