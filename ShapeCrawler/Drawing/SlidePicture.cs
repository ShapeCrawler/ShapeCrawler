using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    /// <inheritdoc cref="IPicture" />
    internal class SlidePicture : SlideShape, IPicture
    {
        private readonly StringValue picReference;

        internal SlidePicture(P.Picture pPicture, SCSlide parentSlide, StringValue picReference)
            : base(pPicture, parentSlide, null)
        {
            this.picReference = picReference;
        }

        public SCImage Image => SCImage.CreatePictureImage(this, this.ParentSlide.SlidePart, this.picReference);
    }
}