using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Services;
using SlideDotNet.Services.Placeholders;
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

        /// <summary>
        /// <inheritdoc cref="IShapeContext.PreSettings"/>
        /// </summary>
        public IPreSettings PreSettings { get; }

        public SlidePlaceholderFontService PlaceholderFontService { get; }

        public OpenXmlCompositeElement XmlElement { get; set; }

        public SlidePart XmlSlidePart { get; }

        #endregion Properties

        #region Constructors

        public ShapeContext(IPreSettings preSettings, SlidePlaceholderFontService fontService, OpenXmlCompositeElement xmlElement, SlidePart xmlSldPart)
        {
            PreSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
            PlaceholderFontService = fontService ?? throw new ArgumentNullException(nameof(fontService));
            XmlElement = xmlElement ?? throw new ArgumentNullException(nameof(xmlElement));
            XmlSlidePart = xmlSldPart ?? throw new ArgumentNullException(nameof(xmlSldPart));
            _masterOtherFonts = new Lazy<Dictionary<int, int>>(InitMasterOtherFonts);
        }

        #endregion Constructors

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

        private Dictionary<int, int> InitMasterOtherFonts()
        {
            var result = FontHeightParser.FromCompositeElement(XmlSlidePart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles.OtherStyle);

            return result;
        }
    }
}
