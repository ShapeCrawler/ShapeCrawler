using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide Master.
    /// </summary>
    public abstract class MasterShape : Shape
    {
        protected MasterShape(OpenXmlCompositeElement pShapeTreeChild, SlideSc slide) : base(pShapeTreeChild, slide)
        {

        }

        public override IPlaceholder Placeholder => MasterPlaceholder.Create(PShapeTreeChild);
    }
}