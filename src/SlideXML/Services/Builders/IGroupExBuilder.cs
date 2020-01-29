using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideXML.Models.Settings;
using SlideXML.Models.SlideComponents;

namespace SlideXML.Services.Builders
{
    /// <summary>
    /// Provides APIs to build instance of the <see cref="GroupSL"/> class.
    /// </summary>
    public interface IGroupExBuilder
    {
        GroupSL Build(GroupShape compositeElement, SlidePart sldPart, IPreSettings preSettings);
    }
}