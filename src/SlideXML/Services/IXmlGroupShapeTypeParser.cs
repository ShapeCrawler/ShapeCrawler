using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Services
{
    /// <summary>
    /// Provides APIs to parse <see cref="GroupShapeType"/> instance.
    /// </summary>
    public interface IXmlGroupShapeTypeParser
    {
        IEnumerable<ElementCandidate> CreateCandidates(P.GroupShapeType groupTypeShape, bool groupParsed = true);
    }
}