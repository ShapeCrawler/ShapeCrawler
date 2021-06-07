using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide Master.
    /// </summary>
    internal abstract class MasterShape : Shape, IPresentationComponent
    {
        protected MasterShape(SCSlideMaster parentSlideMaster, OpenXmlCompositeElement sdkPShapeTreeChild)
            : base(sdkPShapeTreeChild, parentSlideMaster)
        {
            this.ParentSlideMaster = parentSlideMaster;
        }

        public override IPlaceholder Placeholder => MasterPlaceholder.Create(this.SdkPShapeTreeChild);

        public SCPresentation ParentPresentation => this.ParentSlideMaster.ParentPresentation;

        public override SCSlideMaster ParentSlideMaster { get; }
    }
}