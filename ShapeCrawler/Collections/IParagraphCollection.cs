using System.Collections.Generic;

namespace ShapeCrawler.Collections
{
    public interface IParagraphCollection : IReadOnlyList<SCParagraph>
    {
        SCParagraph Add();
    }
}