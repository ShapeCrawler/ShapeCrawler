﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;

namespace ShapeCrawler.Settings
{
    internal class ShapeContext
    {
        #region Constructors

        private ShapeContext()
        {
        }

        #endregion Constructors

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