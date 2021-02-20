using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories;
using ShapeCrawler.Factories.Placeholders;
using ShapeCrawler.Shared;

namespace ShapeCrawler.Settings
{
    public class ShapeContext
    {
        private readonly Lazy<Dictionary<int, int>> _masterOtherStyle;

        #region Constructors

        private ShapeContext()
        {
            _masterOtherStyle = new Lazy<Dictionary<int, int>>(InitMasterOtherStyle);
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        ///     Tries to find matched font height from master/layout slides.
        /// </summary>
        /// <param name="paragraphLvl"></param>
        /// <param name="fontSize"></param>
        internal bool TryGetFromMasterOtherStyle(int paragraphLvl, out int fontSize)
        {
            if (_masterOtherStyle.Value.ContainsKey(paragraphLvl))
            {
                fontSize = _masterOtherStyle.Value[paragraphLvl];
                return true;
            }

            fontSize = -1;
            return false;
        }

        #endregion Public Methods

        #region Private Methods

        private Dictionary<int, int> InitMasterOtherStyle()
        {
            var result =
                FontHeightParser.FromCompositeElement(SlidePart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles
                    .OtherStyle);

            return result;
        }

        #endregion Private Methods

        #region Builder

        internal class Builder
        {
            private readonly PlaceholderFontService _fontService;
            private readonly IPlaceholderService _placeholderService;
            private readonly SlidePart _sdkSldPart;

            #region Public Methods

            internal ShapeContext Build(OpenXmlCompositeElement openXmlElement)
            {
                Check.NotNull(openXmlElement, nameof(openXmlElement));

                return new ShapeContext
                {
                    PlaceholderFontService = _fontService,
                    PlaceholderService = _placeholderService,
                    SlidePart = _sdkSldPart,
                    CompositeElement = openXmlElement
                };
            }

            #endregion Public Methods

            #region Constructors

            public Builder(PlaceholderFontService fontService, SlidePart sdkSldPart) :
                this(fontService, sdkSldPart, new PlaceholderService(sdkSldPart.SlideLayoutPart))
            {
            }

            internal Builder(
                PlaceholderFontService fontService,
                SlidePart sdkSldPart,
                IPlaceholderService placeholderService)
            {
                _fontService = fontService;
                _sdkSldPart = sdkSldPart;
                _placeholderService = placeholderService;
            }

            #endregion Constructors
        }

        #endregion Builder

        #region Internal Properties

        internal SlidePart SlidePart { get; private set; }

        internal OpenXmlCompositeElement CompositeElement { get; private set; }

        internal PlaceholderFontService PlaceholderFontService { get; private set; }

        internal IPlaceholderService PlaceholderService { get; private set; }

        #endregion Internal Properties
    }
}