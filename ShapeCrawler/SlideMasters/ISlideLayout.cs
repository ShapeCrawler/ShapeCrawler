using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Shared;

namespace ShapeCrawler.SlideMasters
{
    /// <summary>
    ///     Represents a Slide Layout.
    /// </summary>
    public interface ISlideLayout : IBaseSlide
    {
        /// <summary>
        ///     Gets parent Slide Master.
        /// </summary>
        ISlideMaster ParentSlideMaster { get; }
    }

    /// <summary>
    ///     Represents a Slide Layout.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — Shape Crawler")]
    internal class SCSlideLayout : ISlideLayout
    {
        private readonly ResettableLazy<ShapeCollection> shapes;
        private readonly SCSlideMaster slideMaster;

        internal SCSlideLayout(SCSlideMaster parentSlideMaster, SlideLayoutPart slideLayoutPart)
        {
            this.slideMaster = parentSlideMaster;
            this.SlideLayoutPart = slideLayoutPart;
            this.shapes = new ResettableLazy<ShapeCollection>(() =>
                ShapeCollection.CreateForSlideLayout(slideLayoutPart.SlideLayout.CommonSlideData.ShapeTree, this));
        }

        public IShapeCollection Shapes => this.shapes.Value;

        public ISlideMaster ParentSlideMaster => this.slideMaster;

        internal SlideLayoutPart SlideLayoutPart { get; }

        public void ThrowIfRemoved()
        {
            throw new System.NotImplementedException();
        }
    }
}