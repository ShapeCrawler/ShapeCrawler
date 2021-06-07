using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
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

        internal SlidePicture(
            SCSlide slide,
            ShapeContext spContext,
            P.Picture pPicture,
            StringValue picReference)
            : base(slide, pPicture)
        {
            this.Context = spContext;
            this.picReference = picReference;
        }

        #region Public Properties

        public SCImage Image => SCImage.GetPictureImage(this, this.ParentSlide.SlidePart, this.picReference);

        #endregion Public Properties

        internal ShapeContext Context { get; }
        public SCPresentation ParentPresentation => this.ParentSlideMaster.ParentPresentation;
    }
}