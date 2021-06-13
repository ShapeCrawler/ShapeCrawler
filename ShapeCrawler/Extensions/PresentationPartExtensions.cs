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
            // Get the collection of slide IDs
            OpenXmlElementList slideIds = prePart.Presentation.SlideIdList.ChildElements;

            SlideId sldId = (SlideId)slideIds[sldIndex];
            string relId = sldId.RelationshipId;

            // Get the specified slide part from the relationship ID
            SlidePart slidePart = (SlidePart)prePart.GetPartById(relId);

            return slidePart;
        }
    }
}