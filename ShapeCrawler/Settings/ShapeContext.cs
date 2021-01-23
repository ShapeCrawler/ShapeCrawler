using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories;
using ShapeCrawler.Factories.Placeholders;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;

namespace ShapeCrawler.Settings
{
    public class ShapeContext
    {
        private readonly Lazy<Dictionary<int, int>> _masterOtherFonts;

        #region Internal Properties

        internal SlidePart SlidePart { get; private set; }

        internal OpenXmlElement OpenXmlElement { get; private set; }

        internal PresentationData PresentationData { get; private set; }

        internal PlaceholderFontService PlaceholderFontService { get; private set; }

        internal IPlaceholderService PlaceholderService { get; private set; }

        #endregion Internal Properties

        #region Constructors

        private ShapeContext()
        {
            _masterOtherFonts = new Lazy<Dictionary<int, int>>(InitMasterOtherFonts);
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Tries to find matched font height from master/layout slides.
        /// </summary>
        /// <param name="prLvl"></param>
        /// <param name="fh"></param>
        /// <returns></returns>
        public bool TryGetFontSize(int prLvl, out int fh)
        {
            if (prLvl < 1 || prLvl > FormatConstants.MaxPrLevel)
            {
                throw new ArgumentOutOfRangeException(nameof(prLvl));
            }

            fh = -1;
            if (_masterOtherFonts.Value.ContainsKey(prLvl))
            {
                fh = _masterOtherFonts.Value[prLvl];
                return true;
            }

            return false;
        }

        #endregion Public Methods

        #region Private Methods

        private Dictionary<int, int> InitMasterOtherFonts()
        {
            var result = FontHeightParser.FromCompositeElement(SlidePart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles.OtherStyle);

            return result;
        }

        #endregion Private Methods

        #region Builder

        public class Builder
        {
            private readonly SlidePart _sdkSldPart;
            private readonly PresentationData _preSettings;
            private readonly PlaceholderFontService _fontService;
            private readonly IPlaceholderService _placeholderService;

            #region Constructors

            public Builder(PresentationData preSettings, PlaceholderFontService fontService, SlidePart sdkSldPart):
                this(preSettings, fontService, sdkSldPart, new PlaceholderService(sdkSldPart.SlideLayoutPart))
            {

            }

            internal Builder(
                PresentationData preSettings, 
                PlaceholderFontService fontService, 
                SlidePart sdkSldPart, 
                IPlaceholderService placeholderService)
            {
                _preSettings = preSettings;
                _fontService = fontService;
                _sdkSldPart = sdkSldPart;
                _placeholderService = placeholderService;
            }

            #endregion Constructors

            #region Public Methods

            internal ShapeContext Build(OpenXmlCompositeElement openXmlElement)
            {
                Check.NotNull(openXmlElement, nameof(openXmlElement));

                return new ShapeContext
                {
                    PresentationData = _preSettings,
                    PlaceholderFontService = _fontService,
                    PlaceholderService = _placeholderService,
                    SlidePart = _sdkSldPart,
                    OpenXmlElement = openXmlElement
                };
            }

            #endregion Public Methods
        }

        #endregion Builder
    }
}
