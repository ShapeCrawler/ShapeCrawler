﻿using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories.Drawing
{
    /// <summary>
    /// <inheritdoc cref="IImageExFactory"/>
    /// </summary>
    public class ImageExFactory : IImageExFactory
    {
        #region Public Methods

        /// <summary>
        /// <inheritdoc cref="IImageExFactory.TryFromSdkSlide"/>
        /// </summary>
        /// <param name="xmlSldPart"></param>
        /// <returns></returns>
        public ImageSc TryFromSdkSlide(SlidePart xmlSldPart)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));

            ImageSc backgroundImage = null;
            var background = xmlSldPart.Slide.CommonSlideData.Background;
            if (background != null)
            {
                var aBlipFill = background.Descendants<A.BlipFill>().SingleOrDefault();
                backgroundImage = TryFromBlipFill(xmlSldPart, aBlipFill);
            }

            return backgroundImage;
        }

        /// <summary>
        /// <inheritdoc cref="IImageExFactory.TryFromSdkShape"/>
        /// </summary>
        public ImageSc TryFromSdkShape(SlidePart xmlSldPart, OpenXmlCompositeElement ce)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            Check.NotNull(ce, nameof(ce));

            var shape = (P.Shape)ce;
            var aBlipFill = shape.ShapeProperties.GetFirstChild<A.BlipFill>();
            ImageSc backgroundImage = TryFromBlipFill(xmlSldPart, aBlipFill);

            return backgroundImage;
        }

        #endregion Public Methods

        #region Private Methods

        private static ImageSc TryFromBlipFill(SlidePart sldPart, A.BlipFill aBlipFill)
        {
            ImageSc backgroundImage = null;
            var blipRelateId = aBlipFill?.Blip?.Embed?.Value; // try to get blip relationship ID
            if (blipRelateId != null)
            {
                backgroundImage = new ImageSc(sldPart, blipRelateId);
            }

            return backgroundImage;
        }

        #endregion Private Methods
    }
}
