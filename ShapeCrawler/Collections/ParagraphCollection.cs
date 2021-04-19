using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable SuggestVarOrType_BuiltInTypes

namespace ShapeCrawler.Texts
{
    /// <summary>
    ///     <inheritdoc cref="IParagraphCollection"/>
    /// </summary>
    internal class ParagraphCollection : IParagraphCollection
    {
        private readonly ResettableLazy<List<SCParagraph>> paragraphs;

        private readonly OpenXmlCompositeElement textBodyCompositeElement;
        private readonly SCTextBox textBox;

        #region Constructors

        internal ParagraphCollection(OpenXmlCompositeElement textBodyCompositeElement, SCTextBox textBox)
        {
            this.textBodyCompositeElement = textBodyCompositeElement;
            this.textBox = textBox;

            paragraphs = new ResettableLazy<List<SCParagraph>>(GetParagraphs);
        }

        #endregion Constructors

        private List<SCParagraph> GetParagraphs() //TODO: return null if text box is empty
        {
            IEnumerable<A.Paragraph> aParagraphs = textBodyCompositeElement.Elements<A.Paragraph>();

            var paragraphs = new List<SCParagraph>(aParagraphs.Count());
            paragraphs.AddRange(aParagraphs.Select(aParagraph => new SCParagraph(aParagraph, textBox)));

            return paragraphs;
        }

        #region Public Methods

        public IEnumerator<IParagraph> GetEnumerator()
        {
            return paragraphs.Value.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IParagraph this[int index] => paragraphs.Value[index];

        public int Count => paragraphs.Value.Count;

        /// <summary>
        ///     Adds a new paragraph in collection.
        /// </summary>
        /// <returns>Added <see cref="SCParagraph" /> instance.</returns>
        public IParagraph Add()
        {
            // Create a new paragraph from the last paragraph and insert at the end
            A.Paragraph lastAParagraph = paragraphs.Value.Last().AParagraph;
            A.Paragraph newAParagraph = (A.Paragraph) lastAParagraph.CloneNode(true);
            lastAParagraph.InsertAfterSelf(newAParagraph);

            var newParagraph = new SCParagraph(newAParagraph, textBox)
            {
                Text = string.Empty
            };

            paragraphs.Reset();

            return newParagraph;
        }

        #endregion Public Methods
    }
}