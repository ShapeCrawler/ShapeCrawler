using System.Diagnostics.CodeAnalysis;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;

namespace ShapeCrawler
{
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600", MessageId = "Elements should be documented", Justification = "It is an internal")]
    internal interface ITextBoxContainer
    {
        IPlaceholder Placeholder { get; }

        IShape Shape { get; }

        void ThrowIfRemoved();
    }
}