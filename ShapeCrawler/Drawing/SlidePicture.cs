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
        internal SlidePicture(
            SCSlide slide,
            string blipRelateId,
            ShapeContext spContext,
            P.Picture pPicture)
            : base(slide, pPicture)
        {
            this.Image = new SCImage(this.ParentSlide.SlidePart, blipRelateId);
            this.Context = spContext;
        }

        #region Public Properties

        public SCImage Image { get; }

        #endregion Public Properties

        internal ShapeContext Context { get; }
    }
}