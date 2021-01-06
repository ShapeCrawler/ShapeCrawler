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
    /// <summary>
    /// <inheritdoc cref="IShapeContext"/>
    /// </summary>
    public class ShapeContext : IShapeContext
    {
        private readonly Lazy<Dictionary<int, int>> _masterOtherFonts;

        #region Properties

        public SlidePart SdkSlidePart { get; private set; }

        public OpenXmlElement SdkElement { get; private set; }

        public IPresentationData presentationData { get; private set; }

        public PlaceholderFontService PlaceholderFontService { get; private set; }

        public IPlaceholderService PlaceholderService { get; private set; }

        #endregion Properties

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
        public bool TryGetFontHeight(int prLvl, out int fh)
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
            var result = FontHeightParser.FromCompositeElement(SdkSlidePart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles.OtherStyle);

            return result;
        }

        #endregion Private Methods

        #region Builder

        public class Builder
        {
            private readonly SlidePart _sdkSldPart;
            private readonly IPresentationData _preSettings;
            private readonly PlaceholderFontService _fontService;
            private readonly IPlaceholderService _placeholderService;

            #region Constructors

            public Builder(IPresentationData preSettings, PlaceholderFontService fontService, SlidePart sdkSldPart):
                this(preSettings, fontService, sdkSldPart, new PlaceholderService(sdkSldPart.SlideLayoutPart))
            {

            }


            public Builder(IPresentationData preSettings, PlaceholderFontService fontService, SlidePart sdkSldPart, IPlaceholderService placeholderService)
            {
                _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
                _fontService = fontService ?? throw new ArgumentNullException(nameof(fontService));
                _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
                _placeholderService = placeholderService;
            }

            #endregion Constructors

            #region Public Methods

            public IShapeContext Build(OpenXmlElement sdkElement)
            {
                Check.NotNull(sdkElement, nameof(sdkElement));

                return new ShapeContext
                {
                    presentationData = _preSettings,
                    PlaceholderFontService = _fontService,
                    PlaceholderService = _placeholderService,
                    SdkSlidePart = _sdkSldPart,
                    SdkElement = sdkElement
                };
            }

            #endregion Public Methods
        }

        #endregion Builder
    }
}
