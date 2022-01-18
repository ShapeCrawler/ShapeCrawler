using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    internal abstract class MasterShape : Shape, IPresentationComponent
    {
        protected MasterShape(OpenXmlCompositeElement pShapeTreesChild, SCSlideMaster parentSlideLayoutMaster)
            : base(pShapeTreesChild, parentSlideLayoutMaster, null)
        {
            this.ParentSlideMaster = parentSlideLayoutMaster;
        }

        public override IPlaceholder Placeholder => MasterPlaceholder.Create(this.PShapeTreesChild);

        public SCPresentation ParentPresentationInternal => this.ParentPresentationInternal;
    }
}