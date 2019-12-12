using DocumentFormat.OpenXml.Presentation;
using PptxXML.Models.Elements;

namespace PptxXML.Models
{
    /// <summary>
    /// Provides APIs to build instance of the <see cref="GroupEx"/> class.
    /// </summary>
    public interface IGroupBuilder
    {
        GroupEx Build(GroupShape compositeElement);
    }
}