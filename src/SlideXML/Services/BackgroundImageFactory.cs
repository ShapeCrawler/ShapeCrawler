using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Utilities;
using SlideXML.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Services
{
    /// <summary>
    /// Provides APIs to parse background images.
    /// </summary>
    public interface IBackgroundImageFactory
    {
        ImageEx CreateBackgroundSlide(SlidePart sldPart);

        ImageEx CreateBackgroundShape(SlidePart sldPart, P.Shape pShape);
    }


    /// <summary>
    /// Represents a background <see cref="ImageEx"/> instance factory.
    /// </summary>
    public class BackgroundImageFactory : IBackgroundImageFactory
    {
        #region Public Methods

        /// <summary>
        /// Create slide background image.
        /// </summary>
        /// <returns><see cref="ImageEx"/> instance or null if slide does not have background image.</returns>
        public ImageEx CreateBackgroundSlide(SlidePart sldPart)
        {
            Check.NotNull(sldPart, nameof(sldPart));

            ImageEx backgroundImage = null;
            var background = sldPart.Slide.CommonSlideData.Background;
            if (background != null)
            {
                var aBlipFill = background.Descendants<A.BlipFill>().SingleOrDefault();
                backgroundImage = FromBlipFill(sldPart, aBlipFill);
            }

            return backgroundImage;
        }

        /// <summary>
        /// Create shape background image.
        /// </summary>
        /// <returns><see cref="ImageEx"/> instance or null if shape does not have background image.</returns>
        public ImageEx CreateBackgroundShape(SlidePart sldPart, P.Shape pShape)
        {
            Check.NotNull(sldPart, nameof(sldPart));
            Check.NotNull(pShape, nameof(pShape));
            
            var aBlipFill = pShape.ShapeProperties.GetFirstChild<A.BlipFill>();
            ImageEx backgroundImage = FromBlipFill(sldPart, aBlipFill);

            return backgroundImage;
        }

        #endregion Public Methods

        #region Private Methods

        private ImageEx FromBlipFill(SlidePart sldPart, A.BlipFill aBlipFill)
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
