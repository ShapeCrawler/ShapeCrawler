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
            this.ParentSlideMaster = slideMaster;
        }

        public override IPlaceholder Placeholder => MasterPlaceholder.Create(this.PShapeTreeChild);

        public override SCPresentation ParentPresentation => ParentSlideMaster.ParentPresentation; // TODO: remove presentation from shape level

        public override SCSlideMaster ParentSlideMaster { get; }

        internal override ThemePart ThemePart => ParentSlideMaster.PSlideMaster.SlideMasterPart.ThemePart;


    }
}