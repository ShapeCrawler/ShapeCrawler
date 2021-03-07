using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Shared;

namespace ShapeCrawler.SlideMaster
{
    /// <summary>
    ///     Represents a Slide Layout.
    /// </summary>
    public class SlideLayoutSc
    {
        private readonly ResettableLazy<ShapeCollection> _shapes;
        private readonly SlideMasterSc _slideMaster;

        internal SlideLayoutSc(SlideMasterSc slideMaster, SlideLayoutPart sldLayoutPart)
        {
            _slideMaster = slideMaster;
            SlideLayoutPart = sldLayoutPart;
            _shapes = new ResettableLazy<ShapeCollection>(() =>
                ShapeCollection.CreateForSlideLayout(sldLayoutPart.SlideLayout.CommonSlideData.ShapeTree, this));
        }

        internal SlideLayoutPart SlideLayoutPart { get; }

        public ShapeCollection Shapes => _shapes.Value;
        public SlideMasterSc SlideMaster => _slideMaster.Presentation.SlideMasters.GetSlideMasterByLayout(this);
    }
}