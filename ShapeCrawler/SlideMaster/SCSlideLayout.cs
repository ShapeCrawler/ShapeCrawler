using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Shared;

namespace ShapeCrawler.SlideMaster
{
    /// <summary>
    ///     Represents a Slide Layout.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class SCSlideLayout
    {
        private readonly ResettableLazy<ShapeCollection> _shapes;
        private readonly SCSlideMaster _slideMaster;
        internal SlideLayoutPart SlideLayoutPart { get; }

        internal SCSlideLayout(SCSlideMaster slideMaster, SlideLayoutPart sldLayoutPart)
        {
            _slideMaster = slideMaster;
            SlideLayoutPart = sldLayoutPart;
            _shapes = new ResettableLazy<ShapeCollection>(() =>
                ShapeCollection.CreateForSlideLayout(sldLayoutPart.SlideLayout.CommonSlideData.ShapeTree, this));
        }

        public ShapeCollection Shapes => _shapes.Value;
        public SCSlideMaster SlideMaster => _slideMaster.Presentation.SlideMasters.GetSlideMasterByLayout(this);
    }
}