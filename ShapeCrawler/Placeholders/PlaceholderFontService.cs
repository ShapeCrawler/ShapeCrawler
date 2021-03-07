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
        #region Public Methods

        #endregion Constructors

        #region Constructors

        public PlaceholderFontService(SlidePart slidePart, IPlaceholderService placeholderService)
        {
        }

        public PlaceholderFontService(SlidePart slidePart)
            : this(slidePart, new PlaceholderService(slidePart.SlideLayoutPart))
        {
        }

        #endregion Constructors
    }
}