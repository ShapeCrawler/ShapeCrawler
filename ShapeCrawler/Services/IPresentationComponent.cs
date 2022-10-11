using System.Diagnostics.CodeAnalysis;

namespace ShapeCrawler.Services
{
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600:Elements should be documented", Justification = "It is an internal member")]
    internal interface IPresentationComponent
    {
        SCPresentation PresentationInternal { get; }
    }
}