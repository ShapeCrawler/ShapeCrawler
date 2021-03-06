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
        private readonly SlideMasterSc _slideMaster;
        private readonly ResettableLazy<ShapeCollection> _shapes;
        internal SlideLayoutPart SlideLayoutPart { get; }

        internal SlideLayoutSc(SlideMasterSc slideMaster, SlideLayoutPart sldLayoutPart)
        {
            _slideMaster = slideMaster;
            SlideLayoutPart = sldLayoutPart;
            _shapes = new ResettableLazy<ShapeCollection>(() =>
                ShapeCollection.CreateForSlideLayout(sldLayoutPart.SlideLayout.CommonSlideData.ShapeTree, this));
        }

        public ShapeCollection Shapes => _shapes.Value;
        public SlideMasterSc SlideMaster => _slideMaster.Presentation.SlideMasters.GetSlideMasterByLayout(this);
    }
}