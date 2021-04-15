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
        protected MasterShape(SCSlideMaster slideMaster, OpenXmlCompositeElement pShapeTreeChild) : base(
            pShapeTreeChild)
        {
            SlideMaster = slideMaster;
        }

        internal override ThemePart ThemePart => SlideMaster.PSlideMaster.SlideMasterPart.ThemePart;

        public override IPlaceholder Placeholder => MasterPlaceholder.Create(PShapeTreeChild);
        public override SCPresentation Presentation => SlideMaster.Presentation;
        public override SCSlideMaster SlideMaster { get; }
    }
}