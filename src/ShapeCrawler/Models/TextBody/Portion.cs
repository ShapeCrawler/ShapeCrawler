using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ShapeCrawler.Collections;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Models.TextBody
{
    /// <summary>
    /// Represents a paragraph portion.
    /// </summary>
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public class Portion
    {
        private readonly A.Text _aText;

        #region Properties

        /// <summary>
        /// Gets or sets text.
        /// </summary>
        public string Text
        {
            get => _aText.Text;
            set => _aText.Text = value;
        }

        public Font Font { get; }
        internal Paragraph Paragraph { get; }

        /// <summary>
        /// Removes the portion from paragraph.
        /// </summary>
        public void Remove()
        {
            Paragraph.Portions.Remove(this);
        }
        
        #endregion Properties

        #region Constructors

        public Portion(A.Text aText, Paragraph paragraph, int fontSize)
        {
            _aText = aText;
            Paragraph = paragraph;
            Font = new Font(aText, fontSize, this);
        }

        #endregion Constructors
    }
}