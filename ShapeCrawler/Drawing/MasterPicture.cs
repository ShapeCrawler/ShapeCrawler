using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing
{
    /// <summary>
    ///     Represents a picture on a Slide Master.
    /// </summary>
    internal class MasterPicture : MasterShape, IShape
    {
        public ShapeType ShapeType => ShapeType.Picture;

        public MasterPicture(P.Picture pPicture, SCSlideMaster slideMasterInternal)
            : base(pPicture, slideMasterInternal)
        {
        }

        public override SCPresentation PresentationInternal { get; }
    }
}