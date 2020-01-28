using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Utilities;
using SlideXML.Extensions;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Services.Placeholders
{
    /// <summary>
    /// Represents a Slide Layout placeholder service.
    /// </summary>
    public class PlaceholderService : IPlaceholderService
    {
        #region Fields

        private const int CustomGeometryCode = 187;
        private List<PlaceholderSL> _placeholders;

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

            // Local function to generate layout and master placeholders
            static List<PlaceholderSL> GetPlaceholders(IEnumerable<OpenXmlCompositeElement> ce)
            {
                var result = new List<PlaceholderSL>(ce.Count());
                foreach (var el in ce)
                {
                    var type = el.GetPlaceholderType();
                    var idx = el.GetPlaceholderIndex();
                    if (type == null && idx == null)
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

                    result.Add(newPh);
                }

                return result;
            }

            // Get OpenXmlCompositeElement instances have P.ShapeProperties
            var layoutElements = sldLtPart.SlideLayout.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>()
                .Where(el => el.Descendants<P.ShapeProperties>().Any());
            var masterElements = sldLtPart.SlideMasterPart.SlideMaster.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>()
                .Where(el => el.Descendants<P.ShapeProperties>().Any());
            var lHolders = GetPlaceholders(layoutElements);
            var mHolders = GetPlaceholders(masterElements);

            foreach (var mElement in mHolders)
            {
                if (lHolders.Any(x => x.Type == mElement.Type || x.Id == mElement.Id))
                {
                    var shape = (P.Shape)mElement.CompositeElement;
                    var dRp = shape.TextBody.ListStyle?.Level1ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>();
                    if (dRp == null)
                    {
                        continue;
                    }

                    var removeEl = lHolders.Single(x => x.Type == mElement.Type || x.Id == mElement.Id);
                    lHolders.Remove(removeEl);
                    lHolders.Add(mElement);
                }
                else
                {
                    lHolders.Add(mElement);
                }
            }

            _placeholders = lHolders;
        }

        #endregion
    }
}
