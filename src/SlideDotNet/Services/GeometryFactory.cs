using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using SlideDotNet.Enums;
using SlideDotNet.Services.Placeholders;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services
{
    /// <summary>
    /// <inheritdoc cref="IGeometryFactory"/>.
    /// </summary>
    public class GeometryFactory : IGeometryFactory
    {
        #region Fields

        private readonly IPlaceholderService _phService;

        #endregion Fields

        #region Constructors

        public GeometryFactory(IPlaceholderService phService)
        {
            _phService = phService ?? throw new ArgumentNullException(nameof(phService));
        }

        #endregion Constructors

        #region Public Methods

        public GeometryType ForShape(P.Shape sdkShape)
        {
            Check.NotNull(sdkShape, nameof(sdkShape));
            return ForCompositeElement(sdkShape, sdkShape.ShapeProperties);
        }

        public GeometryType ForPicture(P.Picture sdkPicture)
        {
            Check.NotNull(sdkPicture, nameof(sdkPicture));
            return ForCompositeElement(sdkPicture, sdkPicture.ShapeProperties);
        }

        #endregion Public Methods

        #region Private Methods

        private GeometryType ForCompositeElement(OpenXmlCompositeElement sdkCompositeElement, P.ShapeProperties spPr)
        {
            var t2D = spPr.Transform2D;
            if (t2D != null)
            {
                var presetGeometry = spPr.GetFirstChild<PresetGeometry>();
                // Placeholder can have transform on the slide, without having geometry
                if (presetGeometry == null)
                {
                    if (spPr.OfType<CustomGeometry>().Any())
                    {
                        return GeometryType.Custom;
                    }
                    return FromLayout();
                }

                var name = presetGeometry.Preset.Value.ToString();
                Enum.TryParse(name, true, out GeometryType geometryType);
                return geometryType;
            }

            return FromLayout();

            GeometryType FromLayout()
            {
                var placeholderLocationData = _phService.TryGetLocation(sdkCompositeElement);
                if (placeholderLocationData == null)
                {
                    return GeometryType.Rectangle;
                }
                return placeholderLocationData.Geometry;
            }
        }

        #endregion Private Methods
    }
}