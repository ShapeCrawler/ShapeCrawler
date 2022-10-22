using System.Diagnostics.CodeAnalysis;

namespace ShapeCrawler.AutoShapes;

[SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600", MessageId = "Elements should be documented", Justification = "It is an internal member.")]
internal interface ITextFrameContainer
{
    Shape Shape { get; }
        
    ITextFrame? TextFrame { get; }
}