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
    /// <inheritdoc cref="IPlaceholderService"/>
    /// </summary>
    public class PlaceholderService : IPlaceholderService
    {
        #region Fields

        private const int CustomGeometryCode = 187;
        private List<PlaceholderSL> _placeholders;
        private readonly SlideLayoutPart _sldLtPart;

        #endregion Fields

        #region Constructors

        public PlaceholderService(SlideLayoutPart sldLtPart)
        {
            Check.NotNull(sldLtPart, nameof(sldLtPart));
            _sldLtPart = sldLtPart;
            Init(_sldLtPart);
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Tries to return the <see cref="PlaceholderSL"/> instance that satisfies type/identifier of specified element or null if no such instance exists.
        /// </summary>
        /// <param name="ce"></param>
        /// <returns></returns>
        public PlaceholderSL TryGet(OpenXmlCompositeElement ce)
        {
            var type = ce.GetPlaceholderType();
            if (type != null)
            {
                return _placeholders.SingleOrDefault(p => p.Type == type);
            }

            var idx = ce.GetPlaceholderIndex();
            return _placeholders.SingleOrDefault(p => p.Id == idx);
        }

        #endregion

        #region Private Methods

        private void Init(SlideLayoutPart sldLtPart)
        {
            // Get OpenXmlCompositeElement instances have P.ShapeProperties
            var layoutElements = sldLtPart.SlideLayout.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>()
                .Where(el => el.Descendants<P.ShapeProperties>().Any());
            var masterElements = sldLtPart.SlideMasterPart.SlideMaster.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>()
                .Where(el => el.Descendants<P.ShapeProperties>().Any());
            var layoutHolders = GetPlaceholders(layoutElements);
            var masterHolders = GetPlaceholders(masterElements);

            // if master placeholder contains level font height, then it becomes a priority than the layout
            foreach (var mHolder in masterHolders)
            {
                if (layoutHolders.Any(x => x.Type == mHolder.Type || x.Id == mHolder.Id))
                {
                    var shape = (P.Shape)mHolder.CompositeElement;
                    var dRp = shape.TextBody.ListStyle?.Level1ParagraphProperties?.GetFirstChild<DefaultRunProperties>();
                    if (dRp == null)
                    {
                        continue;
                    }

                    var removeEl = layoutHolders.Single(x => x.Type == mHolder.Type || x.Id == mHolder.Id);
                    layoutHolders.Remove(removeEl);
                    layoutHolders.Add(mHolder);
                }
                else
                {
                    layoutHolders.Add(mHolder);
                }
            }

            _placeholders = layoutHolders;

        }

        private List<PlaceholderSL> GetPlaceholders(IEnumerable<OpenXmlCompositeElement> ce)
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
                    CompositeElement = el,
                    SlideLayoutPart = _sldLtPart
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

        #endregion
    }
}
