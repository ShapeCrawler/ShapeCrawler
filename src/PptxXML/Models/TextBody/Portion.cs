using ObjectEx.Utilities;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Models.TextBody
{
    /// <summary>
    /// Represents paragraph text portion.
    /// </summary>
    public class Portion
    {
        #region Fields

        #endregion Fields

        #region Properties

        /// <summary>
        /// Returns font height in EMUs.
        /// </summary>
        public int FontHeight { get; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Portion"/> class.
        /// </summary>
        public Portion(int fontHeight)
        {
            Check.IsPositive(fontHeight, nameof(fontHeight));
            FontHeight = fontHeight;
        }

        #endregion Constructors
    }
}