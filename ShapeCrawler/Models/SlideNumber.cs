namespace ShapeCrawler.Models
{
    /// <summary>
    ///     Represents a slide number data.
    /// </summary>
    public class SlideNumber //TODO: do we really need this separate class to define class number?
    {
        #region Constructors

        internal SlideNumber(int sldNum)
        {
            Number = sldNum;
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        ///     Gets or sets slide number.
        /// </summary>
        public int Number { get; }

        #endregion Properties
    }
}