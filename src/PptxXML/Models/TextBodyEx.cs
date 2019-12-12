using System.Collections.Generic;
using ObjectEx.Utilities;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Models
{
    /// <summary>
    /// Represents a text body of the shape.
    /// </summary>
    public class TextBodyEx
    {
        #region Properties

        /// <summary>
        /// Gets paragraphs.
        /// </summary>
        public IList<ParagraphEx> Paragraphs { get; } = new List<ParagraphEx>();

        #endregion Properties

        #region Constructors

        public TextBodyEx(P.TextBody xmlTxtBody)
        {
            Check.NotNull(xmlTxtBody, nameof(xmlTxtBody));
            ParseParagraphs(xmlTxtBody);
        }

        #endregion Constructors

        #region Private Methods

        private void ParseParagraphs(P.TextBody xmlTxtBody)
        {
            foreach (var p in xmlTxtBody.Descendants<A.Paragraph>())
            {
                Paragraphs.Add(new ParagraphEx(p));
            }
        }

        #endregion Private Methods
    }
}
