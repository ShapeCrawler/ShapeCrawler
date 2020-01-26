using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Utilities;
using SlideXML.Extensions;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Services.Placeholders
{
    /// <summary>
    /// Represents a Slide Layout placeholder service.
    /// </summary>
    public class PlaceholderService : IPlaceholderService
    {
        #region Fields

        private const int CustomGeometryCode = 187;
        private readonly List<PlaceholderSL> _placeholders = new List<PlaceholderSL>();

        #endregion Fields

        #region Constructors

        public PlaceholderService(SlideLayoutPart sldLtPart)
        {
            Init(sldLtPart);
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Gets placeholder.
        /// </summary>
        /// <param name="ce"></param>
        /// <returns></returns>
        public PlaceholderSL Get(OpenXmlCompositeElement ce)
        {
            var type = ce.GetPlaceholderType();
            if (type != null)
            {
                return _placeholders.Single(p => p.Type == type);
            }

            var idx = ce.GetPlaceholderIndex();
            return _placeholders.Single(p => p.Id == idx);
        }

        #endregion

        #region Private Methods

        private void Init(SlideLayoutPart sldLtPart)
        {
            Check.NotNull(sldLtPart, nameof(sldLtPart));

            // Get OpenXmlCompositeElement instances have P.ShapeProperties.
            var layoutElements = sldLtPart.SlideLayout.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>()
                .Where(el => el.Descendants<P.ShapeProperties>().Any());
            var masterElements = sldLtPart.SlideMasterPart.SlideMaster.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>()
                .Where(el => el.Descendants<P.ShapeProperties>().Any());

            foreach (var el in layoutElements.Union(masterElements))
            {
                var type = el.GetPlaceholderType();
                var idx = el.GetPlaceholderIndex();
                if (type == null && idx == null)
                {
                    continue;
                }

                if (type != null && _placeholders.Any(p => p.Type == type))
                {
                    continue;
                }

                if (idx != null && _placeholders.Any(p => p.Id == idx))
                {
                    continue;
                }

                var elShapeProperties = el.Descendants<P.ShapeProperties>().Single();
                var t2d = elShapeProperties.Transform2D;
                if (t2d == null)
                {
                    continue;
                }

                // Gets X, Y, W, H and ShapeProperties
                var newPh = new PlaceholderSL
                {
                    X = t2d.Offset.X.Value,
                    Y = t2d.Offset.Y.Value,
                    Width = t2d.Extents.Cx.Value,
                    Height = t2d.Extents.Cy.Value,
                    CompositeElement = el
                };

                // Gets geometry form
                var presetGeometry = elShapeProperties.GetFirstChild<PresetGeometry>();
                if (presetGeometry == null)
                {
                    newPh.GeometryCode = CustomGeometryCode;
                }
                else
                {
                    newPh.GeometryCode = (int)presetGeometry.Preset.Value;
                }

                newPh.Type = type;
                newPh.Id = (int?)idx;
                _placeholders.Add(newPh);
            }
        }

        #endregion
    }
}
