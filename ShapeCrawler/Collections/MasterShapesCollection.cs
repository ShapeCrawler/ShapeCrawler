using System.Collections.Generic;
using ShapeCrawler.Models;

namespace ShapeCrawler.Collections
{
    public class MasterShapesCollection : LibraryCollection<BaseShape>
    {
        public MasterShapesCollection(IEnumerable<BaseShape> paragraphItems) : base(paragraphItems)
        {
        }
    }
}