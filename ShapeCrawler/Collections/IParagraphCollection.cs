using System.Collections.Generic;
using ShapeCrawler.AutoShapes;

namespace ShapeCrawler.Collections
{
    public interface IParagraphCollection : IReadOnlyList<SCParagraph>
    {
        SCParagraph Add();
    }
}