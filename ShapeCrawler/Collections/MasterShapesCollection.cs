using System.Collections.Generic;
using ShapeCrawler.Models;
using ShapeCrawler.Models.Experiment;

namespace ShapeCrawler.Collections
{
    public class MasterShapesCollection : LibraryCollection<BaseShape>
    {
        public MasterShapesCollection(IEnumerable<BaseShape> paragraphItems) : base(paragraphItems)
        {
        }
    }
}