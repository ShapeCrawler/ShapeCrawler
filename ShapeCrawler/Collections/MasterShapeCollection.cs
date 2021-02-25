using System.Collections.Generic;
using ShapeCrawler.Experiment;

namespace ShapeCrawler.Collections
{
    public class MasterShapeCollection : LibraryCollection<BaseShape>
    {
        public MasterShapeCollection(IEnumerable<BaseShape> paragraphItems) : base(paragraphItems)
        {
        }
    }
}