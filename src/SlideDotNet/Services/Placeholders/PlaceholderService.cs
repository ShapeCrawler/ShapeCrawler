using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
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

        private HashSet<PlaceholderLocationData> _phLocations;

        #endregion Fields

        #region Constructors

        public PlaceholderService(SlideLayoutPart sldLtPart)
        {
            Check.NotNull(sldLtPart, nameof(sldLtPart));
            Init(sldLtPart);
        }

        #endregion Constructors

        #region Public Methods

        public PlaceholderLocationData TryGetLocation(OpenXmlCompositeElement sdkCompositeElement)
        {
            Check.NotNull(sdkCompositeElement, nameof(sdkCompositeElement));

            if (!sdkCompositeElement.IsPlaceholder())
            {
                return null;
            }

            var placeholderData = CreatePlaceholderData(sdkCompositeElement);
            var result = _phLocations.FirstOrDefault(p => p.Equals(placeholderData));
            if (result == null && placeholderData.Index != null)
            {
                var idx = placeholderData.Index;
                return _phLocations.FirstOrDefault(p => p.PlaceholderType == PlaceholderType.Body && p.Index == idx);
            }

            return result;
        }

        public PlaceholderType GetPlaceholderType(OpenXmlElement sdkElement)
        {
            var sdkPlaceholder = sdkElement.Descendants<P.PlaceholderShape>().First();

            return GetPlaceholderType(sdkPlaceholder);
        }

        /// <summary>
        /// Gets placeholder data from SDK-element.
        /// </summary>
        /// <param name="sdkElement">Placeholder which is placeholder.</param>
        public PlaceholderData CreatePlaceholderData(OpenXmlElement sdkElement)
        {
            Check.NotNull(sdkElement, nameof(sdkElement));

            var result = new PlaceholderData();
            var ph = sdkElement.Descendants<P.PlaceholderShape>().First();

            // TYPE
            result.PlaceholderType = GetPlaceholderType(ph);

            // INDEX
            if (ph.Index != null)
            {
                result.Index = (int)ph.Index.Value;
            }

            return result;
        }

        public PlaceholderFontData PlaceholderFontDataFromCompositeElement(OpenXmlCompositeElement sdkCompositeElement)
        {
            var placeholderData = CreatePlaceholderData(sdkCompositeElement);

            return new PlaceholderFontData
            {
                PlaceholderType = placeholderData.PlaceholderType,
                Index = placeholderData.Index
            };
        }

        #endregion

        #region Private Methods

        private PlaceholderType GetPlaceholderType(P.PlaceholderShape sdkPlaceholder)
        {
            var phTypeXml = sdkPlaceholder.Type;

            if (phTypeXml == null)
            {
                return PlaceholderType.Custom;
            }
            else
            {
                // Simple title and centered title placeholders were united
                if (phTypeXml == P.PlaceholderValues.Title || phTypeXml == P.PlaceholderValues.CenteredTitle)
                {
                    return PlaceholderType.Title;
                }
                else
                {
                    return Enum.Parse<PlaceholderType>(phTypeXml.Value.ToString());
                }
            }
        }

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
            _phLocations = layoutHolders.ToHashSet(); //TODO: optimize ToHashSet()
        }

        private List<PlaceholderLocationData> GetPlaceholders(IEnumerable<OpenXmlCompositeElement> compositeElements)
        {
            var filtered = Filter(compositeElements);
            var result = new List<PlaceholderLocationData>(filtered.Count());
            foreach (var el in filtered)
            {
                var placeholderData = CreatePlaceholderData(el);
                // avoid duplicate non-custom placeholders
                if (result.Any(p => p.Equals(placeholderData)))
                {
                    continue;
                }

                var spPr = el.Descendants<P.ShapeProperties>().First();
                var t2D = spPr.Transform2D;
                var placeholderLocationData = new PlaceholderLocationData(placeholderData)
                {
                    X = t2D.Offset.X.Value,
                    Y = t2D.Offset.Y.Value,
                    Width = t2D.Extents.Cx.Value,
                    Height = t2D.Extents.Cy.Value
                };

                var presetGeometry = spPr.GetFirstChild<PresetGeometry>();
                if (presetGeometry != null)
                {
                    var name = presetGeometry.Preset.Value.ToString();
                    Enum.TryParse(name, true, out GeometryType geometryType);
                    placeholderLocationData.Geometry = geometryType;
                }

                result.Add(placeholderLocationData);
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
