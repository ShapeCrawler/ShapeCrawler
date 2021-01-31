using ShapeCrawler.Shared;

namespace ShapeCrawler.Models
{
    /// <summary>
    /// Represents a slide number data.
    /// </summary>
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
