using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    internal interface ITextBoxContainer // TODO: what about replace with abstract class?
    {
        SCSlideMaster ParentSlideMaster { get; }

        IPlaceholder Placeholder { get; }

        void ThrowIfRemoved();
        
        IShape Shape { get; }
    }
}