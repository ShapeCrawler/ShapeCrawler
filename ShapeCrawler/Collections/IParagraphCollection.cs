using System.Collections.Generic;
using ShapeCrawler.AutoShapes;

namespace ShapeCrawler.Collections
{
    public interface IParagraphCollection : IReadOnlyList<ParagraphSc>
    {
        ParagraphSc Add();
    }
}