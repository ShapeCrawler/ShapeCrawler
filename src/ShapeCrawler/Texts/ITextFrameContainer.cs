using System.Diagnostics.CodeAnalysis;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.Texts;

[SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600", MessageId = "Elements should be documented", Justification = "It is an internal member.")]
internal interface ITextFrameContainer
{
    SCSlideAutoShape AutoShape { get; }

    ITextFrame? TextFrame { get; }
}