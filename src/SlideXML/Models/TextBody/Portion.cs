using System;
using SlideXML.Exceptions;

namespace SlideXML.Models.TextBody
{
    /// <summary>
    /// Represents a paragraph text portion.
    /// </summary>
    public class Portion
    {
        private readonly int _fontHeight;

        #region Properties

        /// <summary>
        /// Returns font height in EMUs.
        /// </summary>
        public int FontHeight {
            get
            {
                if (_fontHeight == -1)
                {
                    throw new FeatureNotYetImplementedException();
                }

                return _fontHeight;
            }
        }

        public string Text { get; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Portion"/> class.
        /// </summary>
        public Portion(int fontHeight, string text)
        {
            Text = text ?? throw new ArgumentNullException(nameof(text));
            _fontHeight = fontHeight;
        }

        #endregion Constructors
    }
}