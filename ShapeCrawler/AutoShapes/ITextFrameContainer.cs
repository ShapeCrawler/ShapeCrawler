using System.Diagnostics.CodeAnalysis;
using ShapeCrawler.Placeholders;

namespace ShapeCrawler.AutoShapes
{
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600", MessageId = "Elements should be documented", Justification = "It is an internal member.")]
    internal interface ITextFrameContainer // TODO: remove it?
    {
        Shape Shape { get; }
        
        ITextFrame? TextFrame { get; }

        void ThrowIfRemoved();
    }
}