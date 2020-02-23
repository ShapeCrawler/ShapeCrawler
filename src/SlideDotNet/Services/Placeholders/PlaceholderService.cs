using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Enums;
using SlideDotNet.Extensions;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Services.Placeholders
{
    /// <summary>
    /// <inheritdoc cref="IPlaceholderService"/>
    /// </summary>
    public class PlaceholderService : IPlaceholderService
    {
        #region Fields

        private List<PlaceholderLocationData> _placeholders; //TODO: consider use here HashSet

        #endregion Fields

        #region Constructors

        public PlaceholderService(SlideLayoutPart sldLtPart)
        {
            Check.NotNull(sldLtPart, nameof(sldLtPart));
            Init(sldLtPart);
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
        public PlaceholderLocationData TryGet(OpenXmlCompositeElement ce)
        {
            if (!ce.IsPlaceholder())
            {
                return null;
            }

            var phXml = PlaceholderDataFrom(ce);
            if (phXml.PlaceholderType == PlaceholderType.Custom)
            {
                return _placeholders.SingleOrDefault(p => p.Index == phXml.Index);
            }

            return _placeholders.SingleOrDefault(p => p.PlaceholderType == phXml.PlaceholderType);
        }

        /// <summary>
        /// Gets placeholder data.
        /// </summary>
        /// <param name="compositeElement">Placeholder which is placeholder.</param>
        public static PlaceholderData PlaceholderDataFrom(OpenXmlCompositeElement compositeElement)
        {
            var result = new PlaceholderData();
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

            // slide master can contain duplicate
            foreach (var mHolder in masterHolders.Where(mHolder => !layoutHolders.Contains(mHolder)))
            {
                layoutHolders.Add(mHolder);
            }

            _placeholders = layoutHolders;
        }

        private List<PlaceholderLocationData> GetPlaceholders(IEnumerable<OpenXmlCompositeElement> compositeElements)
        {
            var filtered = Filter(compositeElements);
            var result = new List<PlaceholderLocationData>(filtered.Count());
            foreach (var el in filtered)
            {
                var spPr = el.Descendants<P.ShapeProperties>().Single();
                var t2d = spPr.Transform2D;
                var phXml = PlaceholderDataFrom(el);
                var newPhSl = new PlaceholderLocationData(phXml)
                {
                    X = t2d.Offset.X.Value,
                    Y = t2d.Offset.Y.Value,
                    Width = t2d.Extents.Cx.Value,
                    Height = t2d.Extents.Cy.Value
                };

                // avoid duplicate non-custom placeholders
                if (result.Any(p => p.Equals(newPhSl)))
                {
                    continue;
                }

                result.Add(newPhSl);
            }

            return result;
        }

        private static IEnumerable<OpenXmlCompositeElement> Filter(IEnumerable<OpenXmlCompositeElement> compositeElements)
        {
            var filteredList = new List<OpenXmlCompositeElement>();
            var candidates = compositeElements.Where(e => e.IsPlaceholder());
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
