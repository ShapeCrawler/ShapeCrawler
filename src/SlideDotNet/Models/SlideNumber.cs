using SlideDotNet.Shared;

namespace SlideDotNet.Models
{
    /// <summary>
    /// Represents a slide number data.
    /// </summary>
    /// <remarks>TODO: consider to use struct instead class.</remarks>
    public class SlideNumber
    {
        #region Properties

        /// <summary>
        /// Gets or sets slide number.
        /// </summary>
        public int Number { get; set; }

        #endregion Properties

        #region Constructors

        public SlideNumber(int sldNum)
        {
            Check.IsPositive(sldNum, nameof(sldNum));
            Number = sldNum;
        }

        #endregion Constructors
    }
}
