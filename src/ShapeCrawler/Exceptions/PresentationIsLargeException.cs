using System;
using SlideDotNet.Enums;

namespace SlideDotNet.Exceptions
{
    /// <summary>
    /// Thrown when presentation size is more than allowable.
    /// </summary>
    public class PresentationIsLargeException : SlideDotNetException
    {
        #region Constructors

        private PresentationIsLargeException(string message) : base(message, (int)ExceptionCodes.PresentationIsLargeException)
        {

        }

        #endregion Constructors

        /// <summary>
        /// Creates a new instance of the <see cref="PresentationIsLargeException"/> class with specifying max presentation size.
        /// </summary>
        /// <param name="maxSize"></param>
        public static PresentationIsLargeException FromMax(int maxSize)
        {
            var message = ExceptionMessages.PresentationIsLarge.Replace("{0}", maxSize.ToString(), StringComparison.Ordinal);
            return new PresentationIsLargeException(message);
        }
    }
}
