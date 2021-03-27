using System;
using System.Diagnostics.CodeAnalysis;

namespace ShapeCrawler.Shared
{
    /// <summary>
    ///     Represents parameter checker.
    /// </summary>
    public static class Check
    {
        /// <summary>
        ///     Checks whether specified object is not null.
        /// </summary>
        /// <param name="param"></param>
        /// <param name="paramName"></param>
        [SuppressMessage("ReSharper", "InvertIf")]
        public static void NotNull(object param, string paramName)
        {
            if (param == null)
            {
                if (!string.IsNullOrWhiteSpace(paramName))
                {
                    throw new ArgumentNullException(paramName);
                }

                throw new ArgumentNullException(paramName);
            }
        }
    }
}