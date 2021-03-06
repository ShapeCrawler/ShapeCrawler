using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide Master.
    /// </summary>
    public abstract class MasterShape : Shape
    {
        protected MasterShape(SlideMasterSc slideMaster, OpenXmlCompositeElement pShapeTreeChild) : base(
            pShapeTreeChild)
        {
            SlideMaster = slideMaster;
        }

        public override ThemePart ThemePart => SlideMaster.PSlideMaster.SlideMasterPart.ThemePart;

        public override IPlaceholder Placeholder => MasterPlaceholder.Create(PShapeTreeChild);
        public override PresentationSc Presentation => SlideMaster.Presentation;
        public override SlideMasterSc SlideMaster { get; }
    }
}