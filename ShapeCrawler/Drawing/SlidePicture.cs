using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Settings;
using ShapeCrawler.Statics;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <inheritdoc cref="IPicture" />
    internal class SlidePicture : SlideShape, IPicture
    {
        internal ShapeContext Context { get; }

        #region Constructors

        internal SlidePicture(
            SCSlide slide,
            string blipRelateId,
            ShapeContext spContext,
            P.Picture pPicture) : base(slide, pPicture)
        {
            Image = new SCImage(Slide.SlidePart, blipRelateId);
            Context = spContext;
        }

        #endregion Constructors

        #region Public Properties

        public SCImage Image { get; }

        #endregion Properties
    }
}