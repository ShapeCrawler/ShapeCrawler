using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using PptxXML.Models;
using PptxXML.Models.Elements;
using PptxXML.Models.Settings;
using PptxXML.Services.Placeholder;

namespace PptxXML.Services
{
    /// <summary>
    /// Provides APIs to create shape tree's elements.
    /// </summary>
    public interface IElementFactory
    {
        Element CreateGroupsElement(ElementCandidate ec, SlidePart sldPart, IPreSettings preSettings);

        Element CreateRootSldElement(ElementCandidate ec, SlidePart sldPart, IPreSettings preSettings, Dictionary<int, Placeholders.PlaceholderEx> phDic);
    }
}