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
    /// Represents a <see cref="SlideLayoutPart"/> object parser.
    /// </summary>
    public class SlideLayoutPartParser : ISlideLayoutPartParser
    {
        #region Fields

        private const int CustomGeometryCode = 187;

        #endregion Fields

        #region Public Methods

        /// <summary>
        /// Gets placeholder data dictionary.
        /// </summary>
        /// <param name="sldLtPart"></param>
        public Dictionary<int, PlaceholderEx> GetPlaceholderDic(SlideLayoutPart sldLtPart)
        {
            Check.NotNull(sldLtPart, nameof(sldLtPart));

            var resultDic = new Dictionary<int, PlaceholderEx>();

            // Get OpenXmlCompositeElement instances have P.ShapeProperties.
            var layoutElements = sldLtPart.SlideLayout.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>()
                .Where(el => el.Descendants<P.ShapeProperties>().Any());
            var masterElements = sldLtPart.SlideMasterPart.SlideMaster.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>()
                .Where(el => el.Descendants<P.ShapeProperties>().Any());

            foreach (var el in layoutElements.Union(masterElements))
            {
                var placeholderIndex = el.GetPlaceholderIndex();
                if (placeholderIndex == null || resultDic.ContainsKey((int)placeholderIndex))
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
                var placeholderData = new PlaceholderEx
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
                    placeholderData.GeometryCode = CustomGeometryCode;
                }
                else
                {
                    placeholderData.GeometryCode = (int)presetGeometry.Preset.Value;
                }

                // Add in result dictionary
                resultDic.Add((int)placeholderIndex, placeholderData);
            }

            return resultDic;
        }

        #endregion
    }
}
