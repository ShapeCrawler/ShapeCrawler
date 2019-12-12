using System.Linq;
using ObjectEx.Utilities;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Models
{
    /// <summary>
    /// Represents a text paragraph.
    /// </summary>
    public class ParagraphEx
    {
        #region Properties

        /// <summary>
        /// Gets paragraph text string.
        /// </summary>
        public string Text { get; private set; }

        #endregion Properties

        #region Constructors

        public ParagraphEx(A.Paragraph xmlParagraph)
        {
            Check.NotNull(xmlParagraph, nameof(xmlParagraph));

            ParseText(xmlParagraph);
        }

        #endregion Constructors

        #region Private Methods

        private void ParseText(A.Paragraph xmlParagraph)
        {
            Text = xmlParagraph.Descendants<A.Text>().Select(t => t.Text).Aggregate((t1, t2) => t1 + t2);
        }

        #endregion Private Methods
    }
}
