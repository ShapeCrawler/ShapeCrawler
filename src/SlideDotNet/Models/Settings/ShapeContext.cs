using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Services;
using SlideDotNet.Statics;

namespace SlideDotNet.Models.Settings
{
    /// <summary>
    /// <inheritdoc cref="IShapeContext"/>
    /// </summary>
    public class ShapeContext : IShapeContext
    {
        private readonly Lazy<Dictionary<int, int>> _masterOtherFonts;

        #region Properties

        public IPreSettings PreSettings { get; private set; }

        public SlidePart SkdSlidePart { get; private set; }

        public PlaceholderFontService PlaceholderFontService { get; private set; }

        public OpenXmlElement SdkElement { get; private set; }

        #endregion Properties

        #region Constructors

        private ShapeContext()
        {
            _masterOtherFonts = new Lazy<Dictionary<int, int>>(InitMasterOtherFonts);
        }

        #endregion Constructors

        #region Public Methods

        public bool TryFromMasterOther(int prLvl, out int fh)
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
            var result = FontHeightParser.FromCompositeElement(SkdSlidePart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles.OtherStyle);

            return result;
        }

        #endregion Private Methods

        #region Builder

        public class Builder
        {
            private readonly IPreSettings _preSettings;
            private readonly PlaceholderFontService _fontService;
            private readonly SlidePart _sdkSldPart;

            public Builder(IPreSettings preSettings, PlaceholderFontService fontService, SlidePart sdkSldPart)
            {
                _preSettings = preSettings;
                _fontService = fontService;
                _sdkSldPart = sdkSldPart;
            }

            public IShapeContext Build(OpenXmlElement openXmlElement)
            {
                return new ShapeContext
                {
                    PreSettings = _preSettings,
                    PlaceholderFontService = _fontService,
                    SkdSlidePart = _sdkSldPart,
                    SdkElement = openXmlElement
                };
            }
        }

        #endregion Builder
    }
}
