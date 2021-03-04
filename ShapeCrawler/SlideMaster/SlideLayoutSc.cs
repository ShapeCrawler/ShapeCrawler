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
        private readonly PresentationSc _presentation;

        internal SlideLayoutSc(SlideLayoutPart slideLayoutPart, PresentationSc presentation)
        {
            _slideLayoutPart = slideLayoutPart;
            _presentation = presentation;
            _shapes = new ResettableLazy<ShapeCollection>(() =>
                ShapeCollection.CreateForSlideLayout(slideLayoutPart.SlideLayout.CommonSlideData.ShapeTree, this));
        }

        public SlideMasterSc SlideMaster => _presentation.SlideMasters.GetSlideMaster(_slideLayoutPart);
    }
}