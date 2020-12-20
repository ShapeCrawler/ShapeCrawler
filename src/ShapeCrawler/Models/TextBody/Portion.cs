using System;
using SlideDotNet.Shared;

namespace SlideDotNet.Models.TextBody
{
    /// <summary>
    /// Represents a paragraph text portion.
    /// </summary>
    public class Portion
    {

        #region Properties

        /// <summary>
        /// Returns font height in EMUs.
        /// </summary>
        public int FontHeight { get; }

        /// <summary>
        /// Returns portion text.
        /// </summary>
        public string Text { get; }

        #endregion Properties

        #region Constructors

        public Portion(string text, int fontHeight)
        {
            Check.IsPositive(fontHeight, nameof(fontHeight));
            FontHeight = fontHeight;
            Text = text ?? throw new ArgumentNullException(nameof(text));
        }

        #endregion Constructors
    }
}