using System;
using SlideDotNet.Exceptions;
using SlideDotNet.Validation;

namespace SlideDotNet.Models.TextBody
{
    /// <summary>
    /// Represents a paragraph text portion.
    /// </summary>
    public class Portion
    {
        #region Fields

        private readonly int _fontHeight = -1;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Returns font height in EMUs.
        /// </summary>
        public int FontHeight {
            get
            {
                if (_fontHeight != -1)
                {
                    return _fontHeight;
                }

                throw new FeatureNotYetImplementedException();
            }
        }

        /// <summary>
        /// Returns portion text.
        /// </summary>
        public string Text { get; }

        #endregion Properties

        #region Constructors

        public Portion(string text)
        {
            Text = text ?? throw new ArgumentNullException(nameof(text));
        }

        public Portion(string text, int fontHeight) : this(text)
        {
            Check.IsPositive(fontHeight, nameof(fontHeight));
            _fontHeight = fontHeight;
        }

        #endregion Constructors
    }
}