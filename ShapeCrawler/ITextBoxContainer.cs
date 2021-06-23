using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    internal interface ITextBoxContainer // TODO: what about replace with abstract class?
    {
        SCSlideMaster ParentSlideMaster { get; }

        IPlaceholder Placeholder { get; }
    }
}