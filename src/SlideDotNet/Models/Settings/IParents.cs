using System.Collections.Generic;

namespace SlideDotNet.Models.Settings
{
    /// <summary>
    /// Provides presentation setting's APIs.
    /// </summary>
    public interface IParents
    {
        Dictionary<int, int> LlvFontHeights { get; }
    }
}