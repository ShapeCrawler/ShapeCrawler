using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace PptxXML.Services.Placeholder
{
    /// <summary>
    /// Provides APIs to parse an instance of <see cref="SlideLayoutPart"/> class.
    /// </summary>
    public interface ISlideLayoutPartParser
    {
        Dictionary<int, PlaceholderData> GetPlaceholderDic(SlideLayoutPart sldLtPart);
    }
}