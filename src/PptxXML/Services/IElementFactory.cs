using PptxXML.Models.Elements;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Services
{
    /// <summary>
    /// Provides APIs to create shape tree's elements.
    /// </summary>
    public interface IElementFactory
    {
        Element CreateShape(ElementCandidate ec);

        Element CreateChart(ElementCandidate ec);

        Element CreateTable(ElementCandidate ec);

        Element CreatePicture(ElementCandidate ec);
    }
}