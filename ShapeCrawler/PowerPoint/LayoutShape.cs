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
        protected LayoutShape(SCSlideLayout parentSlideLayout, OpenXmlCompositeElement sdkPShapeTreeChild)
            : base(sdkPShapeTreeChild, parentSlideLayout)
        {
            this.ParentSlideLayout = parentSlideLayout;
        }

        public override IPlaceholder Placeholder => LayoutPlaceholder.Create(this.SdkPShapeTreeChild, this);

        public SCPresentation ParentPresentation => ((SCSlideMaster)this.ParentSlideLayout.ParentSlideMaster).ParentPresentation;

        internal SCSlideLayout ParentSlideLayout { get; }
    }
}