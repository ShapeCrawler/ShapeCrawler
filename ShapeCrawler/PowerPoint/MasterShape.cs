using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide Master.
    /// </summary>
    internal abstract class MasterShape : Shape
    {
        protected MasterShape(SCSlideMaster slideMaster, OpenXmlCompositeElement pShapeTreeChild)
            : base(pShapeTreeChild)
        {
            this.SlideMaster = slideMaster;
        }

        public override IPlaceholder Placeholder => MasterPlaceholder.Create(this.PShapeTreeChild);

        public override SCPresentation ParentPresentation => this.SlideMaster.Presentation;

        public override SCSlideMaster SlideMaster { get; }

        internal override ThemePart ThemePart => this.SlideMaster.PSlideMaster.SlideMasterPart.ThemePart;


    }
}