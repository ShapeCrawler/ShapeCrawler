using ShapeCrawler.Placeholders;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a text box container.
    /// </summary>
    internal interface ITextBoxContainer // TODO: what about replace with abstract class?
    {
        SCSlideMaster ParentSlideMaster { get; }
        Placeholder Placeholder { get; }
    }
}