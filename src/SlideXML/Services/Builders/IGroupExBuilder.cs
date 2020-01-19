using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideXML.Models.Elements;
using SlideXML.Models.Settings;

namespace SlideXML.Services.Builders
{
    /// <summary>
    /// Provides APIs to build instance of the <see cref="GroupEx"/> class.
    /// </summary>
    public interface IGroupExBuilder
    {
        GroupEx Build(GroupShape compositeElement, SlidePart sldPart, IPreSettings preSettings);
    }
}