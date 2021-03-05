using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Shared;

namespace ShapeCrawler.SlideMaster
{
    /// <summary>
    ///     Represents a Slide Layout.
    /// </summary>
    public class SlideLayoutSc : SlideSc
    {
        private readonly SlideLayoutPart _slideLayoutPart;
        private readonly SlideMasterSc _slideMaster;
        
        internal SlideLayoutPart SlideLayoutPart { get; }

        public SlideLayoutSc(SlideMasterSc slideMaster, SlideLayoutPart sldLayoutPart)
            : base(slideMaster.Presentation)
        {
            _slideMaster = slideMaster;
            SlideLayoutPart = sldLayoutPart;
            _shapes = new ResettableLazy<ShapeCollection>(() =>
                ShapeCollection.CreateForSlideLayout(sldLayoutPart.SlideLayout.CommonSlideData.ShapeTree, this));
        }

        public SlideMasterSc SlideMaster => _slideMaster.Presentation.SlideMasters.GetSlideMasterByLayout(this);
    }
}