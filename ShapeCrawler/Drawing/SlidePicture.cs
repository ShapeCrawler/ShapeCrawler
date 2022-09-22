using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    /// <inheritdoc cref="IPicture" />
    [SuppressMessage("ReSharper", "SuggestBaseTypeForParameterInConstructor", Justification = "Internal member")]
    internal class SlidePicture : SlideShape, IPicture
    {
        private readonly StringValue picReference;

        internal SlidePicture(P.Picture pPicture, SCSlide slide, StringValue picReference)
            : base(pPicture, slide, null)
        {
            this.picReference = picReference;
        }

        public SCImage Image => SCImage.ForPicture(this, this.Slide.SDKSlidePart, this.picReference);

        public SCShapeType ShapeType => SCShapeType.Picture;
    }
}