using System.Collections.Generic;

namespace ShapeCrawler.Collections
{
    public interface IParagraphCollection : IReadOnlyList<IParagraph>
    {
        IParagraph Add();
        void Remove(IEnumerable<IParagraph> removeParagraphs);
    }
}