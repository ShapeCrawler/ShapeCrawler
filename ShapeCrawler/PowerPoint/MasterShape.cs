using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    internal abstract class MasterShape : Shape, IPresentationComponent
    {
        protected MasterShape(OpenXmlCompositeElement pShapeTreesChild, SCSlideMaster parentSlideMaster)
            : base(pShapeTreesChild, parentSlideMaster, null)
        {
            this.ParentSlideMaster = parentSlideMaster;
        }

        public override IPlaceholder Placeholder => MasterPlaceholder.Create(this.PShapeTreesChild);

        public SCPresentation ParentPresentationInternal => this.ParentPresentationInternal;
    }
}