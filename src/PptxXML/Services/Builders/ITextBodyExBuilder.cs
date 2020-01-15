using PptxXML.Models.Settings;
using PptxXML.Models.TextBody;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Services.Builders
{
    /// <summary>
    /// Provide method to build <see cref="TextBodyEx"/> instance.
    /// </summary>
    public interface ITextBodyExBuilder
    {
        TextBodyEx Build(P.TextBody xmlTxtBody, ShapeSettings spSetting);
    }
}
