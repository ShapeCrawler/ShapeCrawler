using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using objectEx.Extensions;
using PptxXML.Extensions;
using PptxXML.Models;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Services
{
    /// <summary>
    /// Represent <see cref="SlideLayoutPart"/> object parser.
    /// </summary>
    public class SlideLayoutPartParser
    {
        #region Fields

        private const int CustomGeometryCode = 187;

        #endregion

        #region Public Methods

        /// <summary>
        /// Gets placeholder data dictionary.
        /// </summary>
        /// <param name="sldLtPart"></param>
        public Dictionary<int, PlaceholderData> GetPlaceholderDic(SlideLayoutPart sldLtPart)
        {
            sldLtPart.ThrowIfNull(nameof(sldLtPart));

            var resultDic = new Dictionary<int, PlaceholderData>();

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
                var shapePr = el.Descendants<P.ShapeProperties>().Single();
                var t2d = shapePr.Transform2D;
                if (t2d == null)
                {
                    continue;
                }

                // Gets X, Y, W, H and ShapeProperties
                var placeholderData = new PlaceholderData
                {
                    X = t2d.Offset.X.Value,
                    Y = t2d.Offset.Y.Value,
                    Width = t2d.Extents.Cx.Value,
                    Height = t2d.Extents.Cy.Value,
                    ShapeProperties = shapePr
                };

                // Gets geometry form
                var presetGeometry = shapePr.GetFirstChild<PresetGeometry>();
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

        /// <summary>
        /// Gets placeholder data from slide layout/master by index.
        /// </summary>
        /// <returns>Null if not found.</returns>
        public PlaceholderData GetByIndex(SlideLayoutPart sldLtPart, int placeholderIndex)
        {
            throw new NotImplementedException();

            var result = sldLtPart.SlideLayout.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>()
                .FirstOrDefault(el => el.Descendants<P.ShapeProperties>().Any() 
                             && el.GetPlaceholderIndex() != null 
                             && el.GetPlaceholderIndex().Equals(placeholderIndex));
        }

        #endregion

    }
}
