using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideDotNet.Validation;

namespace SlideDotNet.Extensions
{
    /// <summary>
    /// Contains extension methods for <see cref="PresentationPart"/> class object.
    /// </summary>
    public static class PresentationPartExtensions
    {
        /// <summary>
        /// Gets a <see cref="SlidePart"/> instance by slide number.
        /// </summary>
        /// <param name="prePart"></param>
        /// <param name="sldNumber"></param>
        /// <returns></returns>
        public static SlidePart GetSlidePartByNumber(this PresentationPart prePart, int sldNumber)
        {
            Check.IsPositive(sldNumber, nameof(sldNumber));
            var slideIndex = --sldNumber;

            return GetSlidePartByIndex(prePart, slideIndex);
        }

        /// <summary>
        /// Gets a <see cref="SlidePart"/> instance by slide index.
        /// </summary>
        /// <param name="prePart"></param>
        /// <param name="sldIndex"></param>
        /// <returns></returns>
        public static SlidePart GetSlidePartByIndex(this PresentationPart prePart, int sldIndex)
        {
            // Get the collection of slide IDs
            OpenXmlElementList slideIds = prePart.Presentation.SlideIdList.ChildElements;

            string relId = ((SlideId)slideIds[sldIndex]).RelationshipId;

            // Get the specified slide part from the relationship ID
            SlidePart slidePart = (SlidePart)prePart.GetPartById(relId);

            return slidePart;
        }
    }
}
