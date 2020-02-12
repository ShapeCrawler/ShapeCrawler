using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Enums;
using SlideXML.Extensions;
using SlideXML.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideXML.Services.Placeholders
{
    /// <summary>
    /// <inheritdoc cref="IPlaceholderService"/>
    /// </summary>
    public class PlaceholderService : IPlaceholderService
    {
        #region Fields

        private const int CustomGeometryCode = 187;
        private List<PlaceholderData> _placeholders; //TODO: consider use here HashSet
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
        /// Tries to get placeholder from the repository.
        /// </summary>
        /// <param name="ce"></param>
        /// <returns></returns>
        /// <remarks>
        /// Some placeholder on a slide has its location (x/y) and size (width/height) data on the slide.
        /// </remarks>
        public PlaceholderData TryGet(OpenXmlCompositeElement ce)
        {
            if (!ce.IsPlaceholder())
            {
                return null;
            }

            var phXml = GetPlaceholderXML(ce);
            if (phXml.PlaceholderType == PlaceholderType.Custom)
            {
                return _placeholders.SingleOrDefault(p => p.Index == phXml.Index);
            }

            return _placeholders.SingleOrDefault(p => p.Type == phXml.PlaceholderType);
        }

        /// <summary>
        /// Get placeholder XML.
        /// </summary>
        /// <param name="compositeElement">Placeholder which is placeholder.</param>
        public static PlaceholderXML GetPlaceholderXML(OpenXmlCompositeElement compositeElement)
        {
            var result = new PlaceholderXML();
            var ph = compositeElement.Descendants<P.PlaceholderShape>().First();
            var phTypeXml = ph.Type;

            // TYPE
            if (phTypeXml == null)
            {
                result.PlaceholderType = PlaceholderType.Custom;
            }
            else
            {
                // Simple title and centered title placeholders were united
                if (phTypeXml == P.PlaceholderValues.Title || phTypeXml == P.PlaceholderValues.CenteredTitle)
                {
                    result.PlaceholderType = PlaceholderType.Title;
                }
                else
                {
                    result.PlaceholderType = Enum.Parse<PlaceholderType>(phTypeXml.Value.ToString());
                }
            }

            // INDEX
            if (ph.Index != null)
            {
                result.Index = (int)ph.Index.Value;
            }

            return result;
        }

        #endregion

        #region Private Methods

        private void Init(SlideLayoutPart sldLtPart)
        {
            // Get OpenXmlCompositeElement instances have P.ShapeProperties
            var layoutElements = sldLtPart.SlideLayout.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>();
            var masterElements = sldLtPart.SlideMasterPart.SlideMaster.CommonSlideData.ShapeTree.Elements<OpenXmlCompositeElement>();
            var layoutHolders = GetPlaceholders(layoutElements);
            var masterHolders = GetPlaceholders(masterElements);

            // if master placeholder contains level font height, then it becomes a priority than the layout
            foreach (var mHolder in masterHolders)
            {
                if (layoutHolders.Any(p => p.Type == mHolder.Type || p.Index == mHolder.Index))
                {
                    var shape = (P.Shape)mHolder.CompositeElement;
                    var dRp = shape.TextBody.ListStyle?.Level1ParagraphProperties?.GetFirstChild<DefaultRunProperties>();
                    if (dRp == null)
                    {
                        continue;
                    }

                    var removeEl = layoutHolders.Single(x => x.Type == mHolder.Type || x.Index == mHolder.Index);
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

        private List<PlaceholderData> GetPlaceholders(IEnumerable<OpenXmlCompositeElement> compositeElements)
        {
            var filtered = Filter(compositeElements);
            var result = new List<PlaceholderData>(filtered.Count());
            foreach (var el in filtered)
            {
                var spPr = el.Descendants<P.ShapeProperties>().Single();
                var t2d = spPr.Transform2D;
                var phXml = GetPlaceholderXML(el);
                var newPhSl = new PlaceholderData(phXml)
                {
                    X = t2d.Offset.X.Value,
                    Y = t2d.Offset.Y.Value,
                    Width = t2d.Extents.Cx.Value,
                    Height = t2d.Extents.Cy.Value,
                    CompositeElement = el,
                    SlideLayoutPart = _sldLtPart
                };

                // avoid duplicate non-custom placeholders
                if (result.Any(p => p.Equals(newPhSl)))
                {
                    continue;
                }

                // sets geometry form
                var presetGeometry = spPr.GetFirstChild<PresetGeometry>();
                newPhSl.GeometryCode = presetGeometry != null ? (int)presetGeometry.Preset.Value : CustomGeometryCode;
              
                result.Add(newPhSl);
            }

            return result;
        }

        private static IEnumerable<OpenXmlCompositeElement> Filter(IEnumerable<OpenXmlCompositeElement> compositeElements)
        {
            var filteredList = new List<OpenXmlCompositeElement>();
            var candidates = compositeElements.Where(e=>e.IsPlaceholder());
            foreach (var c in candidates)
            {
                var shPr = c.Descendants<P.ShapeProperties>().FirstOrDefault();
                if (shPr?.Transform2D != null)
                {
                    filteredList.Add(c);
                }
            }

            return filteredList;
        }

        #endregion
    }
}
