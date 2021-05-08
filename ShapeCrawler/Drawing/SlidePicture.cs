using ShapeCrawler.Drawing;
using ShapeCrawler.Settings;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <inheritdoc cref="IPicture" />
    internal class SlidePicture : SlideShape, IPicture
    {
        #region Constructors

        internal SlidePicture(
            SCSlide slide,
            string blipRelateId,
            ShapeContext spContext,
            P.Picture pPicture) : base(slide, pPicture)
        {
            Image = new SCImage(ParentSlide.SlidePart, blipRelateId);
            Context = spContext;
        }

        #endregion Constructors

        internal ShapeContext Context { get; }

        #region Public Properties

        public SCImage Image { get; }

        #endregion Properties
    }
}