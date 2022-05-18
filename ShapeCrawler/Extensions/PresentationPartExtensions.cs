using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions
{
    /// <summary>
    ///     Contains extension methods for <see cref="PresentationPart" /> class object.
    /// </summary>
    internal static class PresentationPartExtensions
    {
        /// <summary>
        ///     Gets a <see cref="SlidePart" /> instance by slide index.
        /// </summary>
        public static SlidePart GetSlidePartByIndex(this PresentationPart prePart, int sldIndex)
        {
            var slideIds = prePart.Presentation.SlideIdList.ChildElements;

            var pSldId = (SlideId)slideIds[sldIndex];
            string relId = pSldId.RelationshipId;

            // Get the specified slide part from the relationship ID
            SlidePart slidePart = (SlidePart)prePart.GetPartById(relId);

            return slidePart;
        }
    }
}