using System.Collections.Generic;
using PptxXML.Models.Elements;
using PptxXML.Services.Placeholder;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Services
{
    /// <summary>
    /// Provides APIs to create shape tree's elements.
    /// </summary>
    public interface IElementFactory
    {
        Element CreateGroupsElement(ElementCandidate ec);

        Element CreateRootElement(ElementCandidate ec, Dictionary<int, PlaceholderData> phDic);
    }
}