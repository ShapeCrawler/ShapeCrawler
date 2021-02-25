﻿using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    /// <summary>
    ///     Represents a shape location and size data manager.
    /// </summary>
    public class LocationParser
    {
        #region Fields

        private readonly IPlaceholderService _phService;

        #endregion Fields

        #region Constructors

        internal LocationParser(IPlaceholderService phService)
        {
            _phService = phService ?? throw new ArgumentNullException(nameof(phService));
        }

        #endregion Constructors

        #region Public Methods

        public ILocation FromComposite(OpenXmlCompositeElement sdkCompositeElement)
        {
            Check.NotNull(sdkCompositeElement, nameof(sdkCompositeElement));

            ILocation innerTransform;
            var aTransform = sdkCompositeElement.Descendants<A.Transform2D>().FirstOrDefault();

            if (aTransform != null
                || sdkCompositeElement.Descendants<P.Transform>().FirstOrDefault() != null
            ) // p:graphicFrame contains p:xfrm
            {
                // Group
                if (sdkCompositeElement.Parent is P.GroupShape groupShape)
                {
                    innerTransform = new NonPlaceholderGroupedTransform(sdkCompositeElement, groupShape);
                }
                // ShapeTree
                else
                {
                    innerTransform = new NonPlaceholderTransform(sdkCompositeElement);
                }
            }
            else
            {
                var placeholderLocationData = _phService.TryGetLocation(sdkCompositeElement);
                innerTransform = new PlaceholderTransform(placeholderLocationData);
            }

            return innerTransform;
        }

        #endregion Public Methods
    }
}