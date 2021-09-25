using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    /// <inheritdoc cref="IPicture" />
    internal class SlidePicture : SlideShape, IPicture
    {
        private readonly StringValue picReference;

        internal SlidePicture(SCSlide slide, P.Picture pPicture, StringValue picReference)
            : base(slide, pPicture)
        {
            this.picReference = picReference;
        }

        public SCImage Image => SCImage.CreatePictureImage(this, this.ParentSlide.SlidePart, this.picReference);
    }
}