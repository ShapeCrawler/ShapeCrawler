using SlideDotNet.Validation;

namespace SlideDotNet.Models
{
    public class SlideNumber
    {
        /// <summary>
        /// Gets or sets slide number.
        /// </summary>
        public int Number { get; set; }

        public SlideNumber(int sldNum)
        {
            Check.IsPositive(sldNum, nameof(sldNum));
            Number = sldNum;
        }
    }
}
