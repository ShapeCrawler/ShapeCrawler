using PptxXML.Models.Settings;
using PptxXML.Models.TextBody;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Services.Builders
{
    /// <summary>
    /// Provide method to build <see cref="ParagraphEx.ParagraphExBuilder"/> instance.
    /// </summary>
    public interface IParagraphExBuilder
    {
        ParagraphEx Build(A.Paragraph aParagraph, ShapeSettings spSetting);
    }
}
