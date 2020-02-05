using System.Collections.Generic;

namespace SlideXML.Models.Settings
{
    /// <summary>
    /// Provides presentation setting's APIs.
    /// </summary>
    public interface IPreSettings
    {
        Dictionary<int, int> LlvFontHeights { get; }
    }
}