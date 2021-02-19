using System.Collections.Generic;
using ShapeCrawler.Texts;

namespace ShapeCrawler.Collections
{
    public interface IParagraphCollection : IReadOnlyList<ParagraphSc>
    {
        ParagraphSc Add();
    }
}