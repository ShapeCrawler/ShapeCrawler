using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models;
using SlideDotNet.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services
{
    /// <summary>
    /// <inheritdoc cref="IBackgroundImageFactory"/>
    /// </summary>
    public class BackgroundImageFactory : IBackgroundImageFactory
    {
        #region Public Methods

        /// <summary>
        /// <inheritdoc cref="IBackgroundImageFactory.FromXmlSlide"/>
        /// </summary>
        /// <param name="xmlSldPart"></param>
        /// <returns></returns>
        public ImageEx FromXmlSlide(SlidePart xmlSldPart)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));

            ImageEx backgroundImage = null;
            var background = xmlSldPart.Slide.CommonSlideData.Background;
            if (background != null)
            {
                var aBlipFill = background.Descendants<A.BlipFill>().SingleOrDefault();
                backgroundImage = FromBlipFill(xmlSldPart, aBlipFill);
            }

            return backgroundImage;
        }

        /// <summary>
        /// <inheritdoc cref="IBackgroundImageFactory.FromXmlShape"/>
        /// </summary>
        public ImageEx FromXmlShape(SlidePart xmlSldPart, P.Shape xmlShape)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            Check.NotNull(xmlShape, nameof(xmlShape));
            
            var aBlipFill = xmlShape.ShapeProperties.GetFirstChild<A.BlipFill>();
            ImageEx backgroundImage = FromBlipFill(xmlSldPart, aBlipFill);

            return backgroundImage;
        }

        #endregion Public Methods

        #region Private Methods

        private static ImageEx FromBlipFill(SlidePart sldPart, A.BlipFill aBlipFill)
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
