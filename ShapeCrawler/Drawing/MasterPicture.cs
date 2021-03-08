using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a picture on a Slide Master.
    /// </summary>
    internal class MasterPicture : MasterShape, IShape
    {
        public MasterPicture(SlideMasterSc slideMaster, P.Picture pPicture)
            : base(slideMaster, pPicture)
        {
        }
    }
}