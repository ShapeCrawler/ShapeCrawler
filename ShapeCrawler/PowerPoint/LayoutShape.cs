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
        protected LayoutShape(SCSlideLayout slideLayout, OpenXmlCompositeElement pShapeTreeChild)
            : base(pShapeTreeChild, slideLayout, null)
        {
            this.SlideLayoutInternal = slideLayout;
            this.SlideMasterInternal = (SCSlideMaster)this.SlideLayoutInternal.SlideMaster;
        }

        public override IPlaceholder Placeholder => LayoutPlaceholder.Create(this.PShapeTreesChild, this);

        public override SCPresentation PresentationInternal => ((SCSlideMaster)this.SlideLayoutInternal.SlideMaster).Presentation;

        public SCSlideLayout SlideLayoutInternal { get; }
        
        internal override SCSlideMaster SlideMasterInternal { get; set; }
    }
}