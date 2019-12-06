using PptxXML.Models.Elements;

namespace PptxXML.Services
{
    /// <summary>
    /// Provides APIs to create slide elements.
    /// </summary>
    public interface IElementCreator
    {
        Element CreateShape(ElementCandidate ec);

        Element CreateChart(ElementCandidate elementCandidate);

        Element CreateTable(ElementCandidate elementCandidate);

        Element CreatePicture(ElementCandidate elementCandidate);
    }
}