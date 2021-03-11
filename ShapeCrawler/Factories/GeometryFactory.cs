using System;
using ShapeCrawler.Placeholders;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    public class GeometryFactory
    {
        #region Fields

        private readonly IPlaceholderService _placeholderService;

        #endregion Fields

        #region Constructors

        internal GeometryFactory(IPlaceholderService placeholderService)
        {
            _placeholderService = placeholderService ?? throw new ArgumentNullException(nameof(placeholderService));
        }

        #endregion Constructors
    }
}