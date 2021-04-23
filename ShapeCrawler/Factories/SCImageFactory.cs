using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    internal class SCImageFactory
    {
        #region Private Methods

        private static SCImage TryFromBlipFill(SlidePart slidePart, A.BlipFill aBlipFill)
        {
            SCImage backgroundImage = null;
            var blipRelateId = aBlipFill?.Blip?.Embed?.Value; // try to get blip relationship ID
            if (blipRelateId != null)
            {
                backgroundImage = new SCImage(slidePart, blipRelateId);
            }

            return backgroundImage;
        }

        #endregion Private Methods

        #region Public Methods

        public static SCImage FromSlidePart(SlidePart slidePart)
        {
            SCImage backgroundImage = null;
            P.Background pBackground = slidePart.Slide.CommonSlideData.Background;
            if (pBackground != null)
            {
                A.BlipFill aBlipFill = pBackground.Descendants<A.BlipFill>().SingleOrDefault();
                backgroundImage = TryFromBlipFill(slidePart, aBlipFill);
            }

            return backgroundImage;
        }

        public SCImage FromSlidePart(SlidePart slidePart, OpenXmlCompositeElement compositeElement)
        {
            P.Shape pShape = (P.Shape) compositeElement;
            A.BlipFill aBlipFill = pShape.ShapeProperties.GetFirstChild<A.BlipFill>();
            SCImage backgroundImage = TryFromBlipFill(slidePart, aBlipFill);

            return backgroundImage;
        }

        #endregion Public Methods
    }
}