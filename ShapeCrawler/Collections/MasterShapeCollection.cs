using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Experiment;

namespace ShapeCrawler.Collections
{
    public class MasterShapeCollection : LibraryCollection<BaseShape>
    {
        private List<Shape> slideMasterShapes;

        public MasterShapeCollection(IEnumerable<BaseShape> paragraphItems) : base(paragraphItems)
        {
        }

        public MasterShapeCollection(List<Shape> slideMasterShapes)
        {
            this.slideMasterShapes = slideMasterShapes;
        }

        public Shape GetShapeByPPlaceholderShape(PlaceholderShape pPlaceholderShape)
        {
            throw new System.NotImplementedException();
        }
    }
}