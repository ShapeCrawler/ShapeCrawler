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
        protected LayoutShape(SCSlideLayout parentSlideInternalLayout, OpenXmlCompositeElement sdkPShapeTreeChild)
            : base(sdkPShapeTreeChild, parentSlideInternalLayout, null)
        {
            this.ParentSlideInternalLayout = parentSlideInternalLayout;
        }

        public override IPlaceholder Placeholder => LayoutPlaceholder.Create(this.PShapeTreesChild, this);

        public override SCSlideMaster ParentSlideMaster { get; set; }

        public SCPresentation ParentPresentationInternal => ((SCSlideMaster)this.ParentSlideInternalLayout.ParentSlideMaster).ParentPresentation;

        public SCSlideLayout ParentSlideInternalLayout { get; }
    }
}