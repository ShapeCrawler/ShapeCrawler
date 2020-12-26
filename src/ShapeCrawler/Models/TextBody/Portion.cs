using ShapeCrawler.Collections;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Models.TextBody
{
    /// <summary>
    /// Represents a portion of text inside a text paragraph.
    /// </summary>
    public class Portion
    {
        private readonly A.Text _aText;
        private readonly Paragraph _paragraph;

        #region Properties

        /// <summary>
        /// Gets font height in EMUs.
        /// </summary>
        public int FontHeight { get; }

        /// <summary>
        /// Gets or sets text.
        /// </summary>
        public string Text
        {
            get => _aText.Text;
            set => _aText.Text = value;
        }

        /// <summary>
        /// Removes the portion from paragraph.
        /// </summary>
        public void Remove()
        {
            _paragraph.Portions.Remove(this);
            _aText.Parent.Remove(); // removes from DOM
        }
        
        #endregion Properties

        #region Constructors

        public Portion(A.Text aText, int fontHeight, Paragraph paragraph)
        {
            Check.IsPositive(fontHeight, nameof(fontHeight));
            FontHeight = fontHeight;
            _aText = aText;
            _paragraph = paragraph;
        }

        #endregion Constructors
    }
}