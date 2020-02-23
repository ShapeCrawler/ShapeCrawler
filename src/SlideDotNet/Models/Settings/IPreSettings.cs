using System.Collections.Generic;

namespace SlideDotNet.Models.Settings
{
    /// <summary>
    /// Represents a global presentation settings.
    /// </summary>
    public interface IPreSettings
    {
        /// <summary>
        /// Returns font heights from global presentation or theme settings.
        /// </summary>
        Dictionary<int, int> LlvFontHeights { get; }
    }
}