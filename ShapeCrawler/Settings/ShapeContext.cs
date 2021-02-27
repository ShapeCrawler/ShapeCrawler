using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;

namespace ShapeCrawler.Settings
{
    internal class ShapeContext
    {
        private readonly Lazy<Dictionary<int, FontData>> _masterOtherStyle;

        #region Constructors

        private ShapeContext()
        {
            _masterOtherStyle = new Lazy<Dictionary<int, FontData>>(InitMasterOtherStyle);
        }

        #endregion Constructors

        #region Public Methods

        internal bool TryGetFromMasterOtherStyle(int paragraphLvl, out int fontSize)
        {
            if (_masterOtherStyle.Value.ContainsKey(paragraphLvl))
            {
                if (_masterOtherStyle.Value[paragraphLvl].FontSize != null)
                {
                    fontSize = _masterOtherStyle.Value[paragraphLvl].FontSize;
                    return true;
                }
            }

            fontSize = -1;
            return false;
        }

        #endregion Public Methods

        #region Private Methods

        private Dictionary<int, FontData> InitMasterOtherStyle()
        {
            var result =
                FontDataParser.FromCompositeElement(SlidePart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles
                    .OtherStyle);

            return result;
        }

        #endregion Private Methods

        #region Builder

        internal class Builder
        {
            private readonly PlaceholderFontService _fontService;
            private readonly IPlaceholderService _placeholderService;
            private readonly SlidePart _slidePart;

            #region Constructors

            internal Builder(PlaceholderFontService fontService, SlidePart slidePart)
            {
                _fontService = fontService;
                _slidePart = slidePart;
                _placeholderService = new PlaceholderService(slidePart.SlideLayoutPart);
            }

            #endregion Constructors

            #region Public Methods

            internal ShapeContext Build(OpenXmlCompositeElement compositeElement)
            {
                return new ShapeContext
                {
                    PlaceholderFontService = _fontService,
                    PlaceholderService = _placeholderService,
                    SlidePart = _slidePart,
                    CompositeElement = compositeElement
                };
            }

            #endregion Public Methods
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