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
        protected LayoutShape(SCSlideLayout parentSlideLayout, OpenXmlCompositeElement pShapeTreeChild)
            : base(pShapeTreeChild)
        {
            this.ParentSlideLayout = parentSlideLayout;
        }

        public override IPlaceholder Placeholder => LayoutPlaceholder.Create(this.PShapeTreeChild, this);

        public override SCPresentation ParentPresentation => ((SCSlideMaster)this.ParentSlideLayout.ParentSlideMaster).ParentPresentation;

        internal override ThemePart ThemePart => this.ParentSlideLayout.SlideLayoutPart.SlideMasterPart.ThemePart;

        internal SCSlideLayout ParentSlideLayout { get; }
    }
}