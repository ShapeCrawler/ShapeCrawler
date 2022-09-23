using System.Diagnostics.CodeAnalysis;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.AutoShapes
{
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600", MessageId = "Elements should be documented", Justification = "It is an internal member.")]
    internal interface ITextFrameContainer // TODO: remove it?
    {
        IPlaceholder Placeholder { get; }

        ITextFrame? TextFrame { get; }

        IShape Shape { get; }

        void ThrowIfRemoved();
    }
}