using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;

namespace SlideDotNet.Services.Builders
{
    /// <summary>
    /// Provides APIs to build instance of the <see cref="Group"/> class.
    /// </summary>
    public interface IGroupExBuilder
    {
        Group Build(GroupShape compositeElement, SlidePart sldPart, IParents parents);
    }
}