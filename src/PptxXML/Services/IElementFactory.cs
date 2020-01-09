using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using PptxXML.Models.Elements;
using PptxXML.Services.Placeholder;

namespace PptxXML.Services
{
    /// <summary>
    /// Provides APIs to create shape tree's elements.
    /// </summary>
    public interface IElementFactory
    {
        Element CreateGroupsElement(ElementCandidate ec, SlidePart sldPart);

        Element CreateRootElement(ElementCandidate ec, SlidePart sldPart, Dictionary<int, PlaceholderData> phDic);
    }
}