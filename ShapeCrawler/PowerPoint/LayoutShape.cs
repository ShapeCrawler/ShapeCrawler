using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide Layout.
    /// </summary>
    internal abstract class LayoutShape : Shape
    {
        protected LayoutShape(SCSlideLayout slideLayout, OpenXmlCompositeElement childOfPShapeTree)
            : base(childOfPShapeTree, slideLayout, null)
        {
            this.SlideLayoutInternal = slideLayout;
            this.SlideMasterInternal = (SCSlideMaster)this.SlideLayoutInternal.ParentSlideMaster;
        }

        internal override SCSlideMaster SlideMasterInternal { get; set; }

        public override IPlaceholder Placeholder => LayoutPlaceholder.Create(this.PShapeTreesChild, this);

        public override SCPresentation PresentationInternal => ((SCSlideMaster)this.SlideLayoutInternal.ParentSlideMaster).ParentPresentation;

        public SCSlideLayout SlideLayoutInternal { get; }
    }
}