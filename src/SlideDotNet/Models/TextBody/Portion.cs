using System;

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

        public string Text { get; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Portion"/> class.
        /// </summary>
        public Portion(int fontHeight, string text)
        {
            Text = text ?? throw new ArgumentNullException(nameof(text));
            FontHeight = fontHeight;
        }

        #endregion Constructors
    }
}