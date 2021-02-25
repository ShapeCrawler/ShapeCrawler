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
        private readonly ResettableLazy<List<ParagraphSc>> _paragraphs;

        private readonly OpenXmlCompositeElement _textBodyCompositeElement;
        private readonly TextBoxSc _textBox;

        #region Constructors

        internal ParagraphCollection(OpenXmlCompositeElement textBodyCompositeElement, TextBoxSc textBox)
        {
            _textBodyCompositeElement = textBodyCompositeElement;
            _textBox = textBox;

            _paragraphs = new ResettableLazy<List<ParagraphSc>>(GetParagraphs);
        }

        #endregion Constructors

        private List<ParagraphSc> GetParagraphs()
        {
            // Parse non-empty paragraphs
            IEnumerable<A.Paragraph> nonEmptyAParagraphs = _textBodyCompositeElement.Elements<A.Paragraph>()
                .Where(p => p.Descendants<A.Text>().Any());

            var paragraphs = new List<ParagraphSc>(nonEmptyAParagraphs.Count());
            paragraphs.AddRange(nonEmptyAParagraphs.Select(aParagraph => new ParagraphSc(aParagraph, _textBox)));

            return paragraphs;
        }

        #region Public Methods

        public IEnumerator<ParagraphSc> GetEnumerator()
        {
            return _paragraphs.Value.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public ParagraphSc this[int index] => _paragraphs.Value[index];

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
        /// <returns>Added <see cref="ParagraphSc" /> instance.</returns>
        public ParagraphSc Add()
        {
            // Create a new paragraph from the last paragraph and insert at the end
            A.Paragraph lastAParagraph = _paragraphs.Value.Last().AParagraph;
            A.Paragraph newAParagraph = (A.Paragraph) lastAParagraph.CloneNode(true);
            lastAParagraph.InsertAfterSelf(newAParagraph);

            ParagraphSc newParagraph = new ParagraphSc(newAParagraph, _textBox)
            {
                Text = string.Empty
            };

            _paragraphs.Reset();

            return newParagraph;
        }

        #endregion Public Methods
    }
}