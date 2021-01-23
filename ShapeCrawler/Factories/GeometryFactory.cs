using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Enums;
using ShapeCrawler.Factories.Placeholders;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    public class GeometryFactory
    {
        #region Fields

        private readonly IPlaceholderService _placeholderService;

        #endregion Fields

        #region Constructors

        public GeometryFactory(IPlaceholderService placeholderService)
        {
            _placeholderService = placeholderService ?? throw new ArgumentNullException(nameof(placeholderService));
        }

        #endregion Constructors

        internal GeometryType ForCompositeElement(OpenXmlCompositeElement compositeElement, P.ShapeProperties spPr)
        {
            Transform2D transform2D = spPr.Transform2D;
            if (transform2D != null)
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
                var placeholderLocationData = _placeholderService.TryGetLocation(compositeElement);
                if (placeholderLocationData == null)
                {
                    return GeometryType.Rectangle;
                }
                return placeholderLocationData.Geometry;
            }
        }
    }
}