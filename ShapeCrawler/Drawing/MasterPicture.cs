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
        public MasterPicture(SCSlideMaster slideMaster, P.Picture pPicture)
            : base(slideMaster, pPicture)
        {
        }
    }
}