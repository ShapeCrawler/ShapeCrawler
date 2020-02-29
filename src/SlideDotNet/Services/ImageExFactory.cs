using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models;
using SlideDotNet.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services
{
    /// <summary>
    /// <inheritdoc cref="IImageExFactory"/>
    /// </summary>
    public class ImageExFactory : IImageExFactory
    {
        #region Public Methods

        /// <summary>
        /// <inheritdoc cref="IImageExFactory.TryFromXmlSlide"/>
        /// </summary>
        /// <param name="xmlSldPart"></param>
        /// <returns></returns>
        public ImageEx TryFromXmlSlide(SlidePart xmlSldPart)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));

            ImageEx backgroundImage = null;
            var background = xmlSldPart.Slide.CommonSlideData.Background;
            if (background != null)
            {
                var aBlipFill = background.Descendants<A.BlipFill>().SingleOrDefault();
                backgroundImage = TryFromBlipFill(xmlSldPart, aBlipFill);
            }

            return backgroundImage;
        }

        /// <summary>
        /// <inheritdoc cref="IImageExFactory.TryFromXmlShape"/>
        /// </summary>
        public ImageEx TryFromXmlShape(SlidePart xmlSldPart, OpenXmlCompositeElement ce)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            Check.NotNull(ce, nameof(ce));

            var shape = (P.Shape)ce;
            var aBlipFill = shape.ShapeProperties.GetFirstChild<A.BlipFill>();
            ImageEx backgroundImage = TryFromBlipFill(xmlSldPart, aBlipFill);

            return backgroundImage;
        }

        #endregion Public Methods

        #region Private Methods

        private static ImageEx TryFromBlipFill(SlidePart sldPart, A.BlipFill aBlipFill)
        {
            ImageEx backgroundImage = null;
            var blipRelateId = aBlipFill?.Blip?.Embed?.Value; // try to get blip relationship ID
            if (blipRelateId != null)
            {
                backgroundImage = new ImageEx(sldPart, blipRelateId);
            }

            return backgroundImage;
        }

        #endregion Private Methods
    }
}
