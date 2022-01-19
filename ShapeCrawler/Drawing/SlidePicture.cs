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

        internal SlidePicture(P.Picture pPicture, SCSlide parentSlideLayoutInternal, StringValue picReference)
            : base(pPicture, parentSlideLayoutInternal, null)
        {
            this.picReference = picReference;
        }

        public SCImage Image => SCImage.CreatePictureImage(this, this.ParentSlideLayoutInternal.SlidePart, this.picReference);
    }
}