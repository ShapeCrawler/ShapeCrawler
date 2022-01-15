using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables
{
    /// <summary>
    ///     Represents a table on a Slide Layout.
    /// </summary>
    internal class LayoutTable : LayoutShape, IShape
    {
        internal LayoutTable(SCSlideLayout slideInternalLayout, P.GraphicFrame pGraphicFrame)
            : base(slideInternalLayout, pGraphicFrame)
        {
        }
    }
}