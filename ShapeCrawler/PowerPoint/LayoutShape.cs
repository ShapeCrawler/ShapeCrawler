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
            : base(pShapeTreeChild)
        {
            this.SlideLayout = slideLayout;
        }

        public override IPlaceholder Placeholder => LayoutPlaceholder.Create(this.PShapeTreeChild, this);

        public override SCPresentation ParentPresentation => ((SCSlideMaster)this.SlideLayout.ParentSlideMaster).ParentPresentation;

        //public override SCSlideMaster SlideMaster => this.SlideLayout.ParentSlideMaster;

        internal override ThemePart ThemePart => this.SlideLayout.SlideLayoutPart.SlideMasterPart.ThemePart;

        internal SCSlideLayout SlideLayout { get; }
    }
}