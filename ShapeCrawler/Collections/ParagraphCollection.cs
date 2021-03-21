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
    public class ParagraphCollection : IParagraphCollection
    {
        private readonly ResettableLazy<List<SCParagraph>> _paragraphs;

        private readonly OpenXmlCompositeElement _textBodyCompositeElement;
        private readonly SCTextBox _textBox;

        #region Constructors

        internal ParagraphCollection(OpenXmlCompositeElement textBodyCompositeElement, SCTextBox textBox)
        {
            _textBodyCompositeElement = textBodyCompositeElement;
            _textBox = textBox;

            _paragraphs = new ResettableLazy<List<SCParagraph>>(GetParagraphs);
        }

        #endregion Constructors

        private List<SCParagraph> GetParagraphs()
        {
            IEnumerable<A.Paragraph> aParagraphs = _textBodyCompositeElement.Elements<A.Paragraph>();

            var paragraphs = new List<SCParagraph>(aParagraphs.Count());
            paragraphs.AddRange(aParagraphs.Select(aParagraph => new SCParagraph(aParagraph, _textBox)));

            return paragraphs;
        }

        #region Public Methods

        public IEnumerator<SCParagraph> GetEnumerator()
        {
            return _paragraphs.Value.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public SCParagraph this[int index] => _paragraphs.Value[index];

        public int Count => _paragraphs.Value.Count;

        public void RemoveRange(int index, int removeCount)
        {
            // Remove from outer
            for (int removeIdx = index; removeIdx <= removeCount; removeIdx++)
            {
                _paragraphs.Value[removeIdx].Remove();
            }

            _paragraphs.Reset();
        }

        /// <summary>
        ///     Adds a new paragraph in collection.
        /// </summary>
        /// <returns>Added <see cref="SCParagraph" /> instance.</returns>
        public SCParagraph Add()
        {
            // Create a new paragraph from the last paragraph and insert at the end
            A.Paragraph lastAParagraph = _paragraphs.Value.Last().AParagraph;
            A.Paragraph newAParagraph = (A.Paragraph) lastAParagraph.CloneNode(true);
            lastAParagraph.InsertAfterSelf(newAParagraph);

            SCParagraph newParagraph = new SCParagraph(newAParagraph, _textBox)
            {
                Text = string.Empty
            };

            _paragraphs.Reset();

            return newParagraph;
        }

        #endregion Public Methods

    }
}