using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents paragraph collection.
    /// </summary>
    public interface IParagraphCollection : IReadOnlyList<IParagraph>
    {
        /// <summary>
        ///     Adds a new paragraph in collection.
        /// </summary>
        /// <returns>Added <see cref="SCParagraph" /> instance.</returns>
        IParagraph Add();

        /// <summary>
        ///     Removes specified paragraphs from collection.
        /// </summary>
        void Remove(IEnumerable<IParagraph> removeParagraphs);
    }

    internal class ParagraphCollection : IParagraphCollection
    {
        private readonly ResettableLazy<List<SCParagraph>> paragraphs;
        private readonly TextFrame textBox;

        internal ParagraphCollection(TextFrame textBox)
        {
            this.textBox = textBox;
            this.paragraphs = new ResettableLazy<List<SCParagraph>>(this.GetParagraphs);
        }

        #region Public Properties

        public int Count => this.paragraphs.Value.Count;

        public IParagraph this[int index] => this.paragraphs.Value[index];

        public IEnumerator<IParagraph> GetEnumerator()
        {
            return this.paragraphs.Value.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        #endregion Public Properties

        public IParagraph Add()
        {
            var lastAParagraph = this.paragraphs.Value.Last().AParagraph;
            var newAParagraph = (A.Paragraph)lastAParagraph.CloneNode(true);
            newAParagraph.ParagraphProperties ??= new A.ParagraphProperties();
            lastAParagraph.InsertAfterSelf(newAParagraph);

            var newParagraph = new SCParagraph(newAParagraph, this.textBox)
            {
                Text = string.Empty
            };

            this.paragraphs.Reset();

            return newParagraph;
        }

        public void Remove(IEnumerable<IParagraph> removeParagraphs)
        {
            foreach (var paragraph in removeParagraphs.Cast<SCParagraph>())
            {
                paragraph.AParagraph.Remove();
                paragraph.IsRemoved = true;
            }

            this.paragraphs.Reset();
        }

        private List<SCParagraph> GetParagraphs()
        {
            if (this.textBox.TextBodyElement == null)
            {
                return new List<SCParagraph>(0);
            }

            return this.textBox.TextBodyElement.Elements<A.Paragraph>().Select(aParagraph => new SCParagraph(aParagraph, this.textBox)).ToList();
        }
    }
}