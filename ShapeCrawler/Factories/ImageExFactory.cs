using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    /// <summary>
    ///     <inheritdoc cref="IImageExFactory" />
    /// </summary>
    public class ImageExFactory
    {
        #region Private Methods

        private static SCImage TryFromBlipFill(SlidePart sldPart, A.BlipFill aBlipFill)
        {
            SCImage backgroundImage = null;
            var blipRelateId = aBlipFill?.Blip?.Embed?.Value; // try to get blip relationship ID
            if (blipRelateId != null)
            {
                backgroundImage = new SCImage(sldPart, blipRelateId);
            }

            return backgroundImage;
        }

        #endregion Private Methods

        #region Public Methods

        /// <summary>
        ///     <inheritdoc cref="IImageExFactory.TryFromSdkSlide" />
        /// </summary>
        /// <param name="xmlSldPart"></param>
        /// <returns></returns>
        public SCImage TryFromSdkSlide(SlidePart xmlSldPart)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));

            SCImage backgroundImage = null;
            var background = xmlSldPart.Slide.CommonSlideData.Background;
            if (background != null)
            {
                var aBlipFill = background.Descendants<A.BlipFill>().SingleOrDefault();
                backgroundImage = TryFromBlipFill(xmlSldPart, aBlipFill);
            }

            return backgroundImage;
        }

        /// <summary>
        ///     <inheritdoc cref="IImageExFactory.TryFromSdkShape" />
        /// </summary>
        public SCImage TryFromSdkShape(SlidePart xmlSldPart, OpenXmlCompositeElement ce)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            Check.NotNull(ce, nameof(ce));

            var shape = (P.Shape) ce;
            var aBlipFill = shape.ShapeProperties.GetFirstChild<A.BlipFill>();
            SCImage backgroundImage = TryFromBlipFill(xmlSldPart, aBlipFill);

            return backgroundImage;
        }

        #endregion Public Methods
    }
}