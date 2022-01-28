using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide Layout.
    /// </summary>
    internal abstract class LayoutShape : Shape, IPresentationComponent
    {
        protected LayoutShape(SCSlideLayout parentSlideLayoutInternal, OpenXmlCompositeElement sdkPShapeTreeChild)
            : base(sdkPShapeTreeChild, parentSlideLayoutInternal, null)
        {
            this.ParentSlideLayoutInternal = parentSlideLayoutInternal;
        }

        public override IPlaceholder Placeholder => LayoutPlaceholder.Create(this.PShapeTreesChild, this);

        internal override SCSlideMaster SlideMasterInternal { get; set; }

        public SCPresentation ParentPresentationInternal => ((SCSlideMaster)this.ParentSlideLayoutInternal.ParentSlideMaster).ParentPresentation;

        public SCSlideLayout ParentSlideLayoutInternal { get; }
    }
}