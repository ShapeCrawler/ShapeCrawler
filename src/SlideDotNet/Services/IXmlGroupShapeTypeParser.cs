using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services
{
    /// <summary>
    /// Provides APIs to parse <see cref="GroupShapeType"/> instance.
    /// </summary>
    public interface IXmlGroupShapeTypeParser
    {
        IEnumerable<ElementCandidate> CreateElementCandidates(P.GroupShapeType xmlGroupTypeShape, bool groupParsed = true);
    }
}