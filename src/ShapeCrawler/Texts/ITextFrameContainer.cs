using System.Diagnostics.CodeAnalysis;

namespace ShapeCrawler.Texts;

[SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600", MessageId = "Elements should be documented", Justification = "It is an internal member.")]
internal interface ITextFrameContainer
{
    SCShape SCShape { get; }

    ITextFrame? TextFrame { get; }
}