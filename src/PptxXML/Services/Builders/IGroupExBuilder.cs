using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxXML.Models.Elements;
using PptxXML.Models.Settings;

namespace PptxXML.Services.Builders
{
    /// <summary>
    /// Provides APIs to build instance of the <see cref="GroupEx"/> class.
    /// </summary>
    public interface IGroupExBuilder
    {
        GroupEx Build(GroupShape compositeElement, SlidePart sldPart, IPreSettings preSettings);
    }
}