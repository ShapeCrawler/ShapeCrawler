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
        internal LayoutTable(SCSlideLayout slideLayoutInternal, P.GraphicFrame pGraphicFrame)
            : base(slideLayoutInternal, pGraphicFrame)
        {
        }

        public SCShapeType ShapeType => SCShapeType.Table;
    }
}