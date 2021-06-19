using ShapeCrawler.Placeholders;

namespace ShapeCrawler
{
    internal interface ITextBoxContainer // TODO: what about replace with abstract class?
    {
        SCSlideMaster ParentSlideMaster { get; }

        IPlaceholder Placeholder { get; }
    }
}