using System;
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
    internal class ParagraphCollection : IParagraphCollection
    {
        private readonly ResettableLazy<List<SCParagraph>> paragraphs;
        private readonly OpenXmlCompositeElement textBodyCompositeElement;
        private readonly SCTextBox textBox;

        internal ParagraphCollection(OpenXmlCompositeElement textBodyCompositeElement, SCTextBox textBox)
        {
            this.textBodyCompositeElement = textBodyCompositeElement;
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

        /// <summary>
        ///     Adds a new paragraph in collection.
        /// </summary>
        /// <returns>Added <see cref="SCParagraph" /> instance.</returns>
        public IParagraph Add()
        {
            // Create a new paragraph from the last paragraph and insert at the end
            A.Paragraph lastAParagraph = this.paragraphs.Value.Last().AParagraph;
            A.Paragraph newAParagraph = (A.Paragraph) lastAParagraph.CloneNode(true);
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
            foreach (SCParagraph paragraph in removeParagraphs.Cast<SCParagraph>())
            {
                paragraph.AParagraph.Remove();
                paragraph.IsRemoved = true;
            }

            this.paragraphs.Reset();
        }

        private List<SCParagraph> GetParagraphs() // TODO: return null if text box is empty?
        {
            IEnumerable<A.Paragraph> aParagraphs = this.textBodyCompositeElement.Elements<A.Paragraph>();
            return aParagraphs.Select(aParagraph => new SCParagraph(aParagraph, this.textBox)).ToList();
        }
    }
}